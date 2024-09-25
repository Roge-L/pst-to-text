import pypff
import os
from datetime import datetime
import html2text
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_pst_to_text(pst_file_path, output_folder, max_body_length=10000):
    pst = pypff.file()
    pst.open(pst_file_path)

    def get_header_value(message, header_name):
        try:
            headers = message.get_transport_headers()
            if headers is not None:
                return headers.get(header_name, '')
        except Exception as e:
            logging.warning(f"Error getting header {header_name}: {str(e)}")
        return ''

    def get_body_text(message):
        try:
            body = message.get_plain_text_body()
            if body is None:
                body = message.get_html_body()
                if body is not None:
                    body = html2text.html2text(body)
                else:
                    return "No body text available"
            
            if isinstance(body, bytes):
                body = body.decode('utf-8', errors='replace')
            
            # Truncate body if it's too long
            if len(body) > max_body_length:
                body = body[:max_body_length] + "... [truncated]"
            
            return body
        except Exception as e:
            logging.warning(f"Error getting message body: {str(e)}")
            return "Error retrieving message body"

    def safe_folder_name(name):
        return ''.join(c for c in name if c.isalnum() or c in (' ', '_', '-'))

    def process_folder(folder, path=[]):
        items = []
        try:
            for i in range(folder.get_number_of_sub_messages()):
                try:
                    message = folder.get_sub_message(i)
                    email_data = [
                        f"Subject: {message.get_subject() or 'No subject'}",
                        f"From: {message.get_sender_name() or 'Unknown sender'}",
                        f"To: {get_header_value(message, 'To')}",
                        f"CC: {get_header_value(message, 'Cc')}",
                        f"Date: {message.get_delivery_time().isoformat() if message.get_delivery_time() else 'Unknown'}",
                        f"Folder: {'/'.join(safe_folder_name(p) for p in path)}",
                        "Body:",
                        get_body_text(message),
                        "-" * 50  # Separator between emails
                    ]
                    items.append("\n".join(email_data))
                except Exception as e:
                    logging.error(f"Error processing message: {str(e)}")

            for j in range(folder.get_number_of_sub_folders()):
                try:
                    subfolder = folder.get_sub_folder(j)
                    items.extend(process_folder(subfolder, path + [subfolder.get_name()]))
                except Exception as e:
                    logging.error(f"Error processing subfolder: {str(e)}")
        except Exception as e:
            logging.error(f"Error processing folder: {str(e)}")
        
        return items

    try:
        root_folder = pst.get_root_folder()
        all_items = process_folder(root_folder)

        output_file = f"pst_content_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        if output_folder:
            output_file = os.path.join(output_folder, output_file)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("\n\n".join(all_items))
        
        logging.info(f"Conversion complete. Output file: {output_file}")
    except Exception as e:
        logging.error(f"Error during PST conversion: {str(e)}")
    finally:
        pst.close()

# Usage
pst_file_path = "dist-list.pst"
output_folder = ""
convert_pst_to_text(pst_file_path, output_folder)