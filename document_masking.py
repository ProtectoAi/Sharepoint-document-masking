import os
import time
import logging
import shutil
import argparse
import requests
from msal import ConfidentialClientApplication
from docx import Document
from logging.handlers import TimedRotatingFileHandler
from urllib.parse import urlparse
import json

def configure_logging(log_file_path=None):
    if log_file_path:
        os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
        handler = TimedRotatingFileHandler(log_file_path, when="H", interval=1, backupCount=5)
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logging.basicConfig(level=logging.INFO, handlers=[handler])
    else:
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[logging.StreamHandler()]
        )

logger = logging.getLogger(__name__)

def get_access_token(client_id, client_secret, tenant_id):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]
    app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
    token_result = app.acquire_token_for_client(scopes=scope)
    if "access_token" not in token_result:
        raise Exception(f"Authentication failed: {token_result.get('error_description')}")
    return token_result["access_token"]

def get_site_id(site_hostname, site_path, access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path}"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["id"]

def list_files_in_folder(site_id, folder_name, access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_name}:/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers)
    resp.raise_for_status()
    return resp.json()["value"]

def download_file(site_id, item_id, local_path, access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/content"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers, stream=True)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)
    logger.info(f"Downloaded {local_path}")

def download_word_files_from_sharepoint_graph(site_hostname, site_path, folder_name, local_folder, client_id, client_secret, tenant_id):
    logger.info(f"Authenticating and connecting to SharePoint site {site_hostname}{site_path}")
    access_token = get_access_token(client_id, client_secret, tenant_id)
    site_id = get_site_id(site_hostname, site_path, access_token)
    files = list_files_in_folder(site_id, folder_name, access_token)

    os.makedirs(local_folder, exist_ok=True)
    downloaded_files = []

    for file in files:
        if file["name"].endswith(".docx"):
            local_path = os.path.join(local_folder, file["name"])
            download_file(site_id, file["id"], local_path, access_token)
            downloaded_files.append(local_path)

    return downloaded_files

def validate_mask_api(base_url, auth_key):
    test_payload = {
        "mask": [
            {"value": "Testing - Testing", "format": "Person Name", "token_name": "Text Token"}
        ]
    }
    headers = {
        'Authorization': f'Bearer {auth_key}',
        'Content-Type': 'application/json',
    }
    try:
        response = requests.put(f"{base_url}/mask", headers=headers, json=test_payload)
        response.raise_for_status()
        data = response.json()
        if 'data' in data and isinstance(data['data'], list) and len(data['data']) > 0 and 'token_value' in data['data'][0]:
            logger.info("Mask API validation successful.")
            return True
        else:
            logger.error("Mask API validation failed: Unexpected response format.")
            return False
    except Exception as e:
        logger.error(f"Mask API validation failed: {e}")
        return False

def call_mask_api(base_url, auth_key, mask_payload):
    headers = {
        'Authorization': f'Bearer {auth_key}',
        'Content-Type': 'application/json',
    }
    response = requests.put(f"{base_url}/mask/async", headers=headers, json=mask_payload)
    response.raise_for_status()
    return response.json()

def check_status(base_url, auth_key, tracking_id):
    headers = {
        'Authorization': f'Bearer {auth_key}',
        'Content-Type': 'application/json',
    }
    payload = {
        "status": [{"tracking_id": tracking_id}]
    }
    response = requests.put(f"{base_url}/async-status", headers=headers, json=payload)
    response.raise_for_status()
    return response.json()

def split_text_into_chunks(text, max_words=500):
    words = text.split()
    chunks = []
    for i in range(0, len(words), max_words):
        chunk = " ".join(words[i:i + max_words])
        chunks.append(chunk)
    return chunks

def process_word_files(base_url, auth_key, word_file_paths, output_dir, word_limit=500, archive_dir=None):
    """
    Processes each word file paragraph-wise with masking, splitting paragraphs longer than word_limit into chunks,
    preserving paragraph order and paragraph breaks.
    The word_limit applies per paragraph, not globally.
    """
    for word_path in word_file_paths:
        logger.info(f"Processing Word file: {word_path}")
        try:
            doc = Document(word_path)

            output_txt_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(word_path))[0]}_masked_output.txt")
            os.makedirs(output_dir, exist_ok=True)
            # Clean output file before writing
            open(output_txt_path, 'w', encoding='utf-8').close()

            paragraph_chunks = []  # Will store tuples (original_paragraph_index, chunk_text)

            # Split paragraphs into chunks with max word_limit words per chunk
            for idx, para in enumerate(doc.paragraphs):
                text = para.text.strip()
                if not text:
                    # Preserve blank lines as empty paragraph chunks
                    paragraph_chunks.append((idx, ""))
                    continue

                # Split paragraph into chunks of max word_limit words
                para_chunks = split_text_into_chunks(text, max_words=word_limit)

                # Add these chunks with paragraph index for order preservation
                for chunk in para_chunks:
                    paragraph_chunks.append((idx, chunk))

            if not paragraph_chunks:
                logger.warning(f"No text paragraphs to process in {word_path}")
                continue

            # For each chunk, send mask API request async, collect tracking IDs
            tracking_ids = []
            for idx, (para_idx, chunk_text) in enumerate(paragraph_chunks):
                if not chunk_text.strip():
                    # Empty paragraph chunk: add dummy tracking id to preserve order and write blank line later
                    tracking_ids.append(None)
                    continue
                mask_payload = {"mask": [{"value": chunk_text}]}
                mask_response = call_mask_api(base_url, auth_key, mask_payload)
                tracking_id = mask_response['data'][0]['tracking_id']
                tracking_ids.append(tracking_id)

            # Now poll status for each tracking id and write output preserving paragraph chunk order
            with open(output_txt_path, 'w', encoding='utf-8') as f_out:
                for idx, tracking_id in enumerate(tracking_ids):
                    para_idx, chunk_text = paragraph_chunks[idx]
                    if tracking_id is None:
                        # Empty paragraph chunk - write blank line
                        f_out.write("\n")
                        continue

                    while True:
                        status_response = check_status(base_url, auth_key, tracking_id)
                        logger.info(f"Masking status for tracking ID {tracking_id}: {json.dumps(status_response, indent=2)}")

                        status = status_response['data'][0]['status']
                        if status == 'SUCCESS':
                            result = status_response['data'][0]['result']
                            masked_text = ""
                            for res in result:
                                masked_text += res.get("token_value", "")
                            f_out.write(masked_text.strip() + "\n\n")  # Paragraph break between chunks
                            logger.info(f"Wrote masked chunk {idx+1} (paragraph {para_idx+1})")
                            break
                        elif status in ['IN-PROGRESS', 'PENDING']:
                            logger.info(f"Waiting for masking completion for tracking ID {tracking_id} (status: {status})")
                            time.sleep(3)
                        else:
                            logger.warning(f"Masking failed or unknown status '{status}' for tracking ID {tracking_id}")
                            # fallback: write original chunk text if masking fails
                            f_out.write(chunk_text + "\n\n")
                            break

            if archive_dir:
                os.makedirs(archive_dir, exist_ok=True)
                shutil.move(word_path, os.path.join(archive_dir, os.path.basename(word_path)))
                logger.info(f"Archived {os.path.basename(word_path)} to {archive_dir}")
            else:
                os.remove(word_path)
                logger.info(f"Deleted processed file: {word_path}")

        except Exception as e:
            logger.error(f"Error processing file {word_path}: {e}")

def main(config_path, sharepoint_folder, local_download_dir, output_dir, log_file_path, word_limit, archive_dir=None):
    configure_logging(log_file_path)

    config = {}
    try:
        with open(config_path, 'r') as f:
            section = None
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):  # skip empty or commented lines
                    continue
                if line.startswith('[') and line.endswith(']'):
                    section = line[1:-1]
                    config[section] = {}
                elif section:
                    # Try both ':' and '=' as separators
                    if ':' in line:
                        key, value = line.split(':', 1)
                    elif '=' in line:
                        key, value = line.split('=', 1)
                    else:
                        continue  # Skip lines without valid key-value delimiter
                    config[section][key.strip().strip('"')] = value.strip().strip('"')
    except Exception as e:
        logger.error(f"Error reading config file: {e}")
        return
    if 'protecto' not in config:
        logger.error("Config missing 'protecto' section")
        return

    base_url = config['protecto'].get('BASE_URL')
    auth_key = config['protecto'].get('AUTH_KEY')
    client_id = config['protecto'].get('CLIENT_ID')
    client_secret = config['protecto'].get('CLIENT_SECRET')
    site_url = config['protecto'].get('SITE_URL')
    tenant_id = config['protecto'].get('TENANT_ID')

    if not all([base_url, auth_key, client_id, client_secret, site_url, tenant_id]):
        logger.error("Missing required configuration values")
        return

    if not validate_mask_api(base_url, auth_key):
        logger.error("Mask API validation failed. Exiting.")
        return

    parsed_url = urlparse(site_url)
    site_hostname = parsed_url.netloc
    site_path = parsed_url.path

    word_files = download_word_files_from_sharepoint_graph(
        site_hostname, site_path, sharepoint_folder,
        local_download_dir, client_id, client_secret, tenant_id
    )

    if not word_files:
        logger.warning("No Word files found on SharePoint folder")
        return

    os.makedirs(output_dir, exist_ok=True)
    process_word_files(base_url, auth_key, word_files, output_dir, word_limit, archive_dir)

    logger.info("Word document masking completed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Mask Word documents from SharePoint")
    parser.add_argument("--config_path", required=True, help="Path to config file")
    parser.add_argument("--sharepoint_folder", required=True, help="SharePoint folder name (inside Documents)")
    parser.add_argument("--local_download_dir", required=True, help="Local folder to download Word files")
    parser.add_argument("--output_dir", required=True, help="Directory to write masked output")
    parser.add_argument("--log_file_path", required=False, help="Optional log file path")
    parser.add_argument("--archive_dir", required=False, help="Optional folder to archive processed Word files")
    parser.add_argument("--word_limit", required=False, type=int, default=500, help="Maximum words per paragraph chunk (default 500)")

    args = parser.parse_args()

    main(
        config_path=args.config_path,
        sharepoint_folder=args.sharepoint_folder,
        local_download_dir=args.local_download_dir,
        output_dir=args.output_dir,
        log_file_path=args.log_file_path,
        word_limit=args.word_limit,
        archive_dir=args.archive_dir,
    )
