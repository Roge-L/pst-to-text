import pypff
import os
from datetime import datetime
import html2text
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

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
        
        output_file = os.path.join(output_folder, output_filename)
        
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

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("PST to Text Converter")
        self.pack(fill=tk.BOTH, expand=True)
        self.create_widgets()

    def create_widgets(self):
        self.drop_label = tk.Label(self, text="Drop PST file here or click to select", bg="lightblue", padx=10, pady=10)
        self.drop_label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.drop_label.bind("<Button-1>", self.select_file)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.drop_file)

        self.output_folder_button = tk.Button(self, text="Select Output Folder", command=self.select_output_folder)
        self.output_folder_button.pack(pady=5)

        self.output_folder_label = tk.Label(self, text="Output Folder: Not selected")
        self.output_folder_label.pack(pady=5)

        self.convert_button = tk.Button(self, text="Convert", command=self.start_conversion)
        self.convert_button.pack(pady=10)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="indeterminate")
        self.progress.pack(pady=10)

        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=5)

    def select_file(self, event=None):
        self.pst_file_path = filedialog.askopenfilename(filetypes=[("PST files", "*.pst")])
        if self.pst_file_path:
            self.pst_file_path = os.path.normpath(self.pst_file_path)
            if os.path.exists(self.pst_file_path):
                self.drop_label.config(text=f"Selected: {os.path.basename(self.pst_file_path)}")
            else:
                messagebox.showerror("Error", f"File not found: {self.pst_file_path}")
                self.pst_file_path = None

    def drop_file(self, event):
        self.pst_file_path = event.data
        if self.pst_file_path:
            # Remove the curly braces that TkinterDnD adds
            self.pst_file_path = self.pst_file_path.strip('{}')
            self.pst_file_path = os.path.normpath(self.pst_file_path)
            if os.path.exists(self.pst_file_path):
                self.drop_label.config(text=f"Dropped: {os.path.basename(self.pst_file_path)}")
            else:
                messagebox.showerror("Error", f"File not found: {self.pst_file_path}")
                self.pst_file_path = None

    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
        if self.output_folder:
            self.output_folder = os.path.normpath(self.output_folder)
            self.output_folder_label.config(text=f"Output Folder: {self.output_folder}")

    def start_conversion(self):
        if not hasattr(self, 'pst_file_path') or not self.pst_file_path:
            messagebox.showerror("Error", "Please select a valid PST file first.")
            return
        
        if not os.path.exists(self.pst_file_path):
            messagebox.showerror("Error", f"File not found: {self.pst_file_path}")
            return
        
        if not hasattr(self, 'output_folder') or not self.output_folder:
            messagebox.showerror("Error", "Please select an output folder first.")
            return

        self.progress.start()
        self.status_label.config(text="Converting...")
        self.master.after(100, self.run_conversion)

    def run_conversion(self):
        try:
            convert_pst_to_text(self.pst_file_path, self.output_folder)
            self.status_label.config(text="Conversion completed successfully!")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")
        finally:
            self.progress.stop()

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.geometry("400x300")
    app = Application(master=root)
    app.mainloop()