import pypff
import os
from datetime import datetime
import html2text
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def convert_pst_to_text(pst_file_path, output_folder):
    pst = pypff.file()
    pst.open(pst_file_path)

    def get_header_value(message, header_name):
        try:
            headers = message.get_transport_headers()
            if headers is not None:
                return headers.get(header_name, '')
        except Exception as e:
            logging.debug(f"Error getting header {header_name}: {str(e)}")
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
            
            return body
        except Exception as e:
            logging.debug(f"Error getting message body: {str(e)}")
            return "Error retrieving message body"

    def safe_folder_name(name):
        return ''.join(c for c in name if c.isalnum() or c in (' ', '_', '-'))

    def process_folder(folder, path, output_file, total_processed):
        folder_name = '/'.join(safe_folder_name(p) for p in path)
        total_messages = folder.get_number_of_sub_messages()
        
        if total_messages > 0:
            logging.info(f"Processing folder: {folder_name} ({total_messages} messages)")

        for i in range(total_messages):
            try:
                message = folder.get_sub_message(i)
                email_data = [
                    f"Subject: {message.get_subject() or 'No subject'}",
                    f"From: {message.get_sender_name() or 'Unknown sender'}",
                    f"To: {get_header_value(message, 'To')}",
                    f"CC: {get_header_value(message, 'Cc')}",
                    f"Date: {message.get_delivery_time().isoformat() if message.get_delivery_time() else 'Unknown'}",
                    f"Folder: {folder_name}",
                    "Body:",
                    get_body_text(message),
                    "-" * 50  # Separator between emails
                ]
                
                with open(output_file, 'a', encoding='utf-8') as f:
                    f.write("\n".join(email_data) + "\n\n")
                
                total_processed += 1
                
                if total_processed % 100 == 0:
                    logging.info(f"Processed {total_processed} messages in total")
            
            except Exception as e:
                logging.error(f"Error processing message {i+1}/{total_messages} in {folder_name}: {str(e)}")

        for j in range(folder.get_number_of_sub_folders()):
            try:
                subfolder = folder.get_sub_folder(j)
                total_processed = process_folder(subfolder, path + [subfolder.get_name()], output_file, total_processed)
            except Exception as e:
                logging.error(f"Error processing subfolder: {str(e)}")

        return total_processed

    try:
        root_folder = pst.get_root_folder()
        
        # Generate output filename based on input PST filename
        pst_filename = os.path.basename(pst_file_path)
        pst_name = os.path.splitext(pst_filename)[0]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"{pst_name}_content_{timestamp}.txt"
        
        if output_folder:
            output_file = os.path.join(output_folder, output_filename)
        else:
            output_file = output_filename
        
        # Clear the file if it exists
        open(output_file, 'w').close()
        
        logging.info(f"Starting PST conversion: {pst_file_path}")
        total_processed = process_folder(root_folder, [], output_file, 0)
        
        logging.info(f"Conversion complete. Total messages processed: {total_processed}")
        logging.info(f"Output file: {output_file}")
    except Exception as e:
        logging.error(f"Error during PST conversion: {str(e)}")
    finally:
        pst.close()

# Usage
pst_file_path = "dist-list.pst"
output_folder = ""
convert_pst_to_text(pst_file_path, output_folder)