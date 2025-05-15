import os
import tempfile
import win32print
import win32api
import win32con
import time
import comtypes.client
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import legal
import io
from tkinter import Tk, messagebox, ttk, StringVar, IntVar, Text, Scrollbar, END

class StampPrinter:
    def __init__(self):
        self.root = Tk()
        self.root.title("Stamp Printer Pro")
        self.setup_ui()
        
        # Hardcoded positions (adjust as needed)
        self.serial_x = 268  # X-coordinate from left (in points)
        self.serial_y = 680  # Y-coordinate from bottom (in points)
        
        # COM initialization for Word
        comtypes.CoInitialize()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.mainloop()
    
    def on_close(self):
        comtypes.CoUninitialize()
        self.root.destroy()

    def log_message(self, message):
        """Add message to log area with timestamp"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_area.insert(END, f"[{timestamp}] {message}\n")
        self.log_area.see(END)  # Auto-scroll to bottom
        self.root.update()  # Refresh UI

    def setup_ui(self):
        """Configure the user interface with progress bar and logging"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Input Settings Frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        # Serial Settings
        ttk.Label(settings_frame, text="Starting Serial:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.start_serial = IntVar(value=1)
        ttk.Entry(settings_frame, textvariable=self.start_serial, width=10).grid(row=0, column=1, sticky="w", padx=5)

        ttk.Label(settings_frame, text="Number of Copies:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.copies = IntVar(value=1)
        ttk.Entry(settings_frame, textvariable=self.copies, width=10).grid(row=1, column=1, sticky="w", padx=5)

        # Printer Selection
        printers = self.get_printers()
        self.printer_name = StringVar(value=printers[0] if printers else "")
        ttk.Label(settings_frame, text="Printer:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        ttk.Combobox(settings_frame, textvariable=self.printer_name, values=printers, state="readonly").grid(row=2, column=1, sticky="ew", padx=5)

        # Paper Size
        self.paper_size = StringVar(value="legal")
        ttk.Label(settings_frame, text="Paper Size:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        ttk.Combobox(settings_frame, textvariable=self.paper_size, values=["legal", "letter", "a4"], state="readonly").grid(row=3, column=1, sticky="w", padx=5)

        # Progress Bar
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=1, column=0, pady=10, sticky="ew")

        # Print Button
        ttk.Button(main_frame, text="Print Stamps", command=self.process_stamps).grid(row=2, column=0, pady=10)

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=3, column=0, sticky="nsew", pady=5)
        
        # Scrollable Text Area
        self.log_area = Text(log_frame, height=10, width=60, wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_area.yview)
        self.log_area.configure(yscrollcommand=scrollbar.set)
        
        self.log_area.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

    def get_printers(self):
        """Hardcode specific printer + PDF option"""
        return [
            "Hewlett-Packard HP LaserJet P4014",  # Your exact printer name
            "Microsoft Print to PDF"
        ]

    def convert_word_to_pdf(self, docx_path):
        """Convert Word to PDF using COM"""
        try:
            self.log_message(f"Converting {os.path.basename(docx_path)} to PDF...")
            word = comtypes.client.CreateObject("Word.Application")
            doc = word.Documents.Open(docx_path)
            temp_pdf = os.path.join(tempfile.gettempdir(), "temp_stamp.pdf")
            doc.SaveAs(temp_pdf, FileFormat=17)
            doc.Close()
            word.Quit()
            self.log_message("Conversion successful")
            return temp_pdf
        except Exception as e:
            error_msg = f"Word to PDF conversion failed: {str(e)}"
            self.log_message(error_msg)
            raise Exception(error_msg)

    def add_serial_to_pdf(self, input_pdf, output_pdf, serial):
        """Add serial number to PDF with Arial size 12"""
        try:
            self.log_message(f"Adding serial number {serial:05d}...")
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=legal)
            
            # Set font to Arial size 12
            try:
                can.setFont("Arial", 12)  # Try Arial first
            except:
                # Fallback to Helvetica if Arial not available
                can.setFont("Helvetica", 12)
                self.log_message("Warning: Arial font not available, using Helvetica")
            
            can.drawString(self.serial_x, self.serial_y, f"{serial:05d}")
            can.save()
            packet.seek(0)
            
            original = PdfReader(input_pdf)
            overlay = PdfReader(packet)
            output = PdfWriter()
            
            page = original.pages[0]
            page.merge_page(overlay.pages[0])
            output.add_page(page)
            
            with open(output_pdf, "wb") as f:
                output.write(f)
            self.log_message(f"Stamp {serial:05d} generated successfully")
        except Exception as e:
            error_msg = f"Failed to add serial number: {str(e)}"
            self.log_message(error_msg)
            raise Exception(error_msg)

    def print_pdf(self, pdf_path, serial):
        """Handles both physical and PDF printing"""
        printer_name = self.printer_name.get()
        
        try:
            # PDF Printer Handling (simpler path)
            if "Microsoft Print to PDF" in printer_name:
                output_pdf = os.path.join(
                    os.path.dirname(__file__), 
                    f"stamp_{serial:05d}.pdf"
                )
                self.log_message(f"Saving PDF to {output_pdf}")
                
                # Method 1: Direct copy (most reliable)
                try:
                    import shutil
                    shutil.copyfile(pdf_path, output_pdf)
                    return True
                except Exception as copy_error:
                    self.log_message(f"Direct copy failed: {str(copy_error)}")
                    
                    # Method 2: ShellExecute fallback
                    result = win32api.ShellExecute(
                        0,
                        "printto",
                        pdf_path,
                        f'"/d:Microsoft Print to PDF" /f:"{output_pdf}"',
                        ".",
                        0
                    )
                    if result <= 32:
                        raise Exception(f"PDF save failed (error {result})")
                    return True

            # Physical Printer Handling
            else:
                printer_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
                hprinter = win32print.OpenPrinter(printer_name, printer_defaults)
                
                # Get and modify printer settings
                printer_info = win32print.GetPrinter(hprinter, 2)
                if not printer_info.get("pDevMode"):
                    raise Exception("Printer configuration unavailable")
                    
                # Set paper size
                paper_map = {
                    "legal": win32con.DMPAPER_LEGAL,
                    "letter": win32con.DMPAPER_LETTER, 
                    "a4": win32con.DMPAPER_A4
                }
                printer_info["pDevMode"].PaperSize = paper_map.get(
                    self.paper_size.get().lower(),
                    win32con.DMPAPER_LETTER
                )
                
                # Apply settings and print
                win32print.SetPrinter(hprinter, 2, printer_info, 0)
                win32api.ShellExecute(0, "print", pdf_path, f'"/d:{printer_name}"', ".", 0)
                return True
                
        except Exception as e:
            error_msg = f"Print failed: {str(e)}"
            self.log_message(error_msg)
            raise Exception(error_msg)
            
        finally:
            if 'hprinter' in locals():
                win32print.ClosePrinter(hprinter)

    def monitor_printer(self, printer_name, timeout=120):
        """Wait for print queue to clear"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                hprinter = win32print.OpenPrinter(printer_name)
                jobs = list(win32print.EnumJobs(hprinter, 0, -1, 1))
                win32print.ClosePrinter(hprinter)
                if not jobs:
                    self.log_message("Print job completed successfully")
                    return True
                time.sleep(2)
            except Exception as e:
                self.log_message(f"Printer monitoring error: {str(e)}")
                time.sleep(1)
        return False

    def process_stamps(self):
        """Main processing workflow"""
        try:
            # Clear previous logs
            self.log_area.delete(1.0, END)
            
            # Auto-locate stamp.docx
            docx_path = os.path.join(os.path.dirname(__file__), "stamp.docx")
            if not os.path.exists(docx_path):
                error_msg = "Error: stamp.docx not found in program directory"
                self.log_message(error_msg)
                messagebox.showerror("Error", error_msg)
                return
            
            total_copies = self.copies.get()
            self.progress["maximum"] = total_copies
            self.progress["value"] = 0
            
            self.log_message(f"Starting batch of {total_copies} stamps")
            
            for i in range(total_copies):
                serial = self.start_serial.get() + i
                self.log_message(f"\nProcessing stamp {serial:05d} ({i+1}/{total_copies})")
                
                temp_pdf = os.path.join(tempfile.gettempdir(), f"stamp_{serial:05d}.pdf")
                
                try:
                    # Update progress
                    self.progress["value"] = i + 1
                    self.root.update()
                    
                    # Convert and process
                    pdf_path = self.convert_word_to_pdf(docx_path)
                    self.add_serial_to_pdf(pdf_path, temp_pdf, serial)
                    
                    if not self.print_pdf(temp_pdf, serial):
                        self.log_message("Warning: Printer issue - continuing")
                    
                    # Cleanup
                    os.remove(temp_pdf)
                    os.remove(pdf_path)
                    time.sleep(0.5)  # Brief pause
                    
                except Exception as e:
                    self.log_message(f"Error processing stamp {serial:05d}: {str(e)}")
                    continue
            
            completion_msg = f"Successfully processed {total_copies} stamps"
            self.log_message(f"\n{completion_msg}")
            messagebox.showinfo("Complete", completion_msg)
            self.progress["value"] = 0
            
        except Exception as e:
            error_msg = f"Fatal error: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)
            self.progress["value"] = 0

if __name__ == "__main__":
    StampPrinter()