import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk, ImageOps
import os
import subprocess
import sys
import io
import fitz  # PyMuPDF


def extract_label_from_pdf(pdf_path, target_dpi=203, debug=True):
    """
    Extract a shipping label from a PDF that may be:
    - Letter size with label on one side
    - Rotated 90 or 270 degrees
    - Already the correct 4x6 size

    Returns a PIL Image of the normalized label.
    """
    doc = fitz.open(pdf_path)
    page = doc[0]

    # Get page dimensions in inches
    page_width_in = page.rect.width / 72
    page_height_in = page.rect.height / 72

    if debug:
        print(f"PDF page size: {page_width_in:.2f} x {page_height_in:.2f} inches")

    # Check if this is approximately letter size (8.5 x 11)
    is_letter_size = (
        (7.5 < page_width_in < 9.5 and 10 < page_height_in < 12) or
        (10 < page_width_in < 12 and 7.5 < page_height_in < 9.5)
    )

    if debug:
        print(f"Is letter size: {is_letter_size}")

    # Render at target DPI
    mat = fitz.Matrix(target_dpi / 72, target_dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("ppm")
    img = Image.open(io.BytesIO(img_data))

    doc.close()

    if not is_letter_size:
        # Assume it's already the right size, just return it
        return img

    # For letter size, find the actual label content
    # Convert to grayscale for analysis
    gray = img.convert('L')

    # Use a stricter threshold to find actual printed content (not light artifacts)
    threshold = 240
    def threshold_fn(x):
        return 0 if x < threshold else 255

    binary = gray.point(threshold_fn, mode='L')

    # Find bounding box of dark content
    bbox = ImageOps.invert(binary).getbbox()

    if bbox is None:
        if debug:
            print("No content bounding box found, returning original")
        return img

    if debug:
        print(f"Initial content bbox: {bbox}")
        print(f"Original image size: {img.size}")

    # Check if bbox covers the whole page (common issue with PDFs that have invisible elements)
    bbox_covers_page = (
        bbox[0] < img.width * 0.05 and
        bbox[1] < img.height * 0.05 and
        bbox[2] > img.width * 0.95 and
        bbox[3] > img.height * 0.95
    )

    if bbox_covers_page:
        if debug:
            print("Bbox covers whole page - using column density analysis")

        # Analyze vertical strips to find where the label actually is
        # Shipping labels are typically on the left or right half
        strip_width = img.width // 20
        strip_densities = []

        for i in range(20):
            strip_left = i * strip_width
            strip_right = (i + 1) * strip_width
            strip = binary.crop((strip_left, 0, strip_right, img.height))
            # Count dark pixels (0 = dark in our binary)
            dark_pixels = sum(1 for p in strip.getdata() if p == 0)
            density = dark_pixels / (strip_width * img.height)
            strip_densities.append((i, density))

        if debug:
            print(f"Strip densities: {[(i, f'{d:.3f}') for i, d in strip_densities]}")

        # Find contiguous region with high density (the label)
        # Label typically has >5% density, empty area <1%
        label_strips = [i for i, d in strip_densities if d > 0.02]

        if label_strips:
            left_strip = min(label_strips)
            right_strip = max(label_strips)

            # Convert back to pixel coordinates with some margin
            margin = strip_width
            left = max(0, left_strip * strip_width - margin)
            right = min(img.width, (right_strip + 1) * strip_width + margin)

            # Now find vertical bounds within this horizontal region
            cropped_region = binary.crop((left, 0, right, img.height))
            inverted_region = ImageOps.invert(cropped_region)
            region_bbox = inverted_region.getbbox()

            if region_bbox:
                top = max(0, region_bbox[1] - margin)
                bottom = min(img.height, region_bbox[3] + margin)
                bbox = (left, top, right, bottom)

                if debug:
                    print(f"Refined bbox from density analysis: {bbox}")

    if debug:
        print(f"Final content bbox: {bbox}")

    # Calculate the content dimensions in inches
    content_width_in = (bbox[2] - bbox[0]) / target_dpi
    content_height_in = (bbox[3] - bbox[1]) / target_dpi

    if debug:
        print(f"Content size in inches: {content_width_in:.2f} x {content_height_in:.2f}")

    # Determine the expected label size based on content dimensions
    # Standard sizes: 4x6, 4x2, etc. - find the best fit
    # The content bbox is the printed area - the actual label might be larger

    # For a rotated 4x6 label on letter paper:
    # - Width would be ~6" (the long edge), Height would be ~4" (short edge)
    # For a rotated 4x2 label:
    # - Width would be ~4", Height would be ~2"

    # Expand the crop region to include the full label (with white margins)
    # Use the larger dimension to determine label size
    max_content_dim = max(content_width_in, content_height_in)
    min_content_dim = min(content_width_in, content_height_in)

    # Determine likely label size
    if max_content_dim > 5:  # Likely a 6" edge -> 4x6 label
        label_long_edge = 6.0
        label_short_edge = 4.0
    elif max_content_dim > 3:  # Likely a 4" edge -> 4x2 or 4x3 label
        label_long_edge = 4.0
        label_short_edge = 2.0
    else:  # Small label, just use content bounds
        label_long_edge = max_content_dim + 0.2
        label_short_edge = min_content_dim + 0.2

    if debug:
        print(f"Detected label size: {label_short_edge} x {label_long_edge} inches")

    # Calculate crop region centered on the content bbox but sized for full label
    content_center_x = (bbox[0] + bbox[2]) / 2
    content_center_y = (bbox[1] + bbox[3]) / 2

    # The label is rotated, so long edge is horizontal on the page
    crop_half_width = (label_long_edge * target_dpi) / 2
    crop_half_height = (label_short_edge * target_dpi) / 2

    left = int(max(0, content_center_x - crop_half_width))
    right = int(min(img.width, content_center_x + crop_half_width))
    top = int(max(0, content_center_y - crop_half_height))
    bottom = int(min(img.height, content_center_y + crop_half_height))

    if debug:
        print(f"Crop region for full label: ({left}, {top}, {right}, {bottom})")
        print(f"Crop size: {right - left} x {bottom - top} pixels")

    # Crop to content
    cropped = img.crop((left, top, right, bottom))

    if debug:
        print(f"Cropped size: {cropped.size}")

    # Determine if rotation is needed based on:
    # 1. The cropped dimensions compared to standard 4x6 label
    # 2. Whether content height spans most of the original page height
    crop_width, crop_height = cropped.size
    aspect_ratio = crop_width / crop_height

    # Standard 4x6 label at 203 DPI is about 812 x 1218 pixels
    # Aspect ratio should be ~0.67 (width/height)
    standard_label_ratio = 4.0 / 6.0  # ~0.67

    # Check if the content spans most of the page height (indicating rotated label)
    content_height = bbox[3] - bbox[1]
    page_height = img.height
    spans_page_height = content_height > (page_height * 0.7)

    if debug:
        print(f"Aspect ratio: {aspect_ratio:.2f} (standard 4x6 is {standard_label_ratio:.2f})")
        print(f"Content spans {content_height/page_height*100:.0f}% of page height")
        print(f"Spans most of page height: {spans_page_height}")

    # Rotation logic:
    # - If content is landscape (wider than tall), rotate
    # - If content spans most of page height but is narrow, it's a rotated label
    needs_rotation = False

    if aspect_ratio > 1.0:
        # Clearly landscape - needs rotation
        needs_rotation = True
        if debug:
            print("Landscape orientation detected")
    elif spans_page_height and aspect_ratio > 0.5:
        # Content is tall (spans page) but aspect ratio suggests it's rotated
        # A 4x6 label rotated 90° on letter paper would be ~6" tall and ~4" wide
        # which gives aspect ratio of ~0.67, but on letter it spans ~6/11 = 55% height
        # This eBay label spans most of the height, so it's rotated
        needs_rotation = True
        if debug:
            print("Tall content spanning page height - likely rotated label")

    if needs_rotation:
        if debug:
            print("Rotating 90 degrees counter-clockwise")
        # Rotate 90° counter-clockwise for labels with text reading bottom-to-top
        cropped = cropped.rotate(90, expand=True)

    if debug:
        print(f"Final size: {cropped.size}")

    return cropped


class ZPLPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ZPL Image Printer")
        self.root.geometry("600x850")  # Increased height for all controls
        self.root.resizable(True, True)
        
        self.image_path = None
        self.pil_image = None
        
        # Default printer name
        self.printer_name = "Zebra_Technologies_ZTC_GX420d_2"
        
        # Create widgets
        self.create_widgets()
        
    def create_widgets(self):
        # Image selection frame
        file_frame = ttk.LabelFrame(self.root, text="Image Selection")
        file_frame.pack(fill="x", padx=10, pady=10)
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        
        self.browse_button = ttk.Button(file_frame, text="Browse...", command=self.browse_file)
        self.browse_button.pack(side="right", padx=10, pady=10)
        
        # Image preview frame - with appropriate aspect ratio
        preview_frame = ttk.LabelFrame(self.root, text="Image Preview")
        preview_frame.pack(fill="both", padx=10, pady=10)
        # Set dimensions that respect a 4x6 aspect ratio
        preview_frame.configure(height=300)  # Increased height
        preview_frame.configure(width=500)   # Width that gives roughly 4:6 ratio
        preview_frame.pack_propagate(False)  # Prevent the frame from resizing to fit its contents
        
        self.preview_label = ttk.Label(preview_frame, text="No image selected")
        self.preview_label.pack(padx=10, pady=10, fill="both")
        
        # Settings frame
        settings_frame = ttk.LabelFrame(self.root, text="Print Settings")
        settings_frame.pack(fill="x", padx=10, pady=10)
        
        # Print speed - use only Spinbox for integer values
        ttk.Label(settings_frame, text="Print Speed (0-14):").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.speed_var = tk.IntVar(value=2)
        self.speed_spin = ttk.Spinbox(settings_frame, from_=0, to=14, textvariable=self.speed_var, width=4)
        self.speed_spin.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Darkness - use only Spinbox for integer values
        ttk.Label(settings_frame, text="Darkness (0-30):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.darkness_var = tk.IntVar(value=25)
        self.darkness_spin = ttk.Spinbox(settings_frame, from_=0, to=30, textvariable=self.darkness_var, width=4)
        self.darkness_spin.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        # Quantity - use only Spinbox for integer values
        ttk.Label(settings_frame, text="Quantity (1-99):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.quantity_var = tk.IntVar(value=1)
        self.quantity_spin = ttk.Spinbox(settings_frame, from_=1, to=99, textvariable=self.quantity_var, width=4)
        self.quantity_spin.grid(row=2, column=1, padx=10, pady=5, sticky="w")
        
        # Position X - use only Spinbox for integer values
        ttk.Label(settings_frame, text="Position X:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.pos_x_var = tk.IntVar(value=0)
        self.pos_x_spin = ttk.Spinbox(settings_frame, from_=0, to=800, textvariable=self.pos_x_var, width=4)
        self.pos_x_spin.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        
        # Position Y - use only Spinbox for integer values
        ttk.Label(settings_frame, text="Position Y:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.pos_y_var = tk.IntVar(value=10)
        self.pos_y_spin = ttk.Spinbox(settings_frame, from_=0, to=800, textvariable=self.pos_y_var, width=4)
        self.pos_y_spin.grid(row=4, column=1, padx=10, pady=5, sticky="w")
        
        # Threshold - use only Spinbox for integer values
        ttk.Label(settings_frame, text="B/W Threshold (0-255):").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.threshold_var = tk.IntVar(value=128)
        self.threshold_spin = ttk.Spinbox(settings_frame, from_=0, to=255, textvariable=self.threshold_var, width=4)
        self.threshold_spin.grid(row=5, column=1, padx=10, pady=5, sticky="w")
        
        # just above “Printer:” row
        ttk.Label(settings_frame, text="Label Size:").grid(row=6, column=0, padx=10, pady=5, sticky="w")

        self.label_sizes = {          # dots at 203 dpi
            "4 x 6 in": (812, 1218),
            "4 x 2 in": (812, 406),
        }
        self.label_size_var = tk.StringVar(value="4 x 6 in")
        ttk.Combobox(
            settings_frame,
            textvariable=self.label_size_var,
            values=list(self.label_sizes.keys()),
            state="readonly",
            width=8
        ).grid(row=6, column=1, padx=10, pady=5, sticky="w")

        # Extract label checkbox (for letter-size PDFs with rotated labels)
        self.extract_label_var = tk.BooleanVar(value=True)
        self.extract_label_check = ttk.Checkbutton(
            settings_frame,
            text="Auto-extract label from letter-size PDF",
            variable=self.extract_label_var,
            command=self.on_extract_toggle
        )
        self.extract_label_check.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        # Printer selection
        ttk.Label(settings_frame, text="Printer:").grid(row=8, column=0, padx=10, pady=5, sticky="w")
        self.printer_var = tk.StringVar(value=self.printer_name)
        self.printer_entry = ttk.Entry(settings_frame, textvariable=self.printer_var, width=40)
        self.printer_entry.grid(row=8, column=1, padx=10, pady=5, sticky="we", columnspan=2)
        
        # Make the columns expand properly
        settings_frame.columnconfigure(1, weight=1)
        
        # Action buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        self.print_button = ttk.Button(button_frame, text="Print", command=self.print_image, state="disabled")
        self.print_button.pack(side="right", padx=5)
        
        # Status frame
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", side="bottom", padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="w")
        status_label.pack(fill="x")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Image or PDF File",
            filetypes=(
                ("Image and PDF files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.pdf"),
                ("Image files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            )
        )
        
        if file_path:
            self.image_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.load_preview()
            self.print_button.config(state="normal")
            self.status_var.set(f"Selected: {os.path.basename(file_path)}")

    def on_extract_toggle(self):
        """Reload preview when extract option is toggled."""
        if self.image_path:
            self.load_preview()

    def load_preview(self):
        try:
            # Check if the file is a PDF
            if self.image_path.lower().endswith('.pdf'):
                pdf_document = fitz.open(self.image_path)
                if len(pdf_document) == 0:
                    raise ValueError("PDF file is empty")
                page_count = len(pdf_document)
                pdf_document.close()

                if self.extract_label_var.get():
                    # Use the smart extraction function
                    self.pil_image = extract_label_from_pdf(self.image_path, target_dpi=203)
                    # Show info about extraction
                    if page_count > 1:
                        self.status_var.set(f"PDF has {page_count} pages - extracted label from page 1")
                    else:
                        self.status_var.set(f"Label extracted: {self.pil_image.size[0]}x{self.pil_image.size[1]} px")
                else:
                    # Render PDF without extraction (original behavior)
                    pdf_document = fitz.open(self.image_path)
                    page = pdf_document[0]
                    mat = fitz.Matrix(203/72, 203/72)
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("ppm")
                    self.pil_image = Image.open(io.BytesIO(img_data))
                    pdf_document.close()
                    if page_count > 1:
                        self.status_var.set(f"PDF has {page_count} pages - using page 1")
            else:
                # Open the image normally
                self.pil_image = Image.open(self.image_path)

            # Get the preview frame size
            preview_width = 480  # Slightly less than the frame width to account for padding
            preview_height = 280  # Slightly less than the frame height to account for padding

            # Calculate aspect ratio
            width, height = self.pil_image.size
            ratio = min(preview_width / width, preview_height / height)
            new_size = (int(width * ratio), int(height * ratio))

            preview_image = self.pil_image.copy()
            preview_image.thumbnail(new_size, Image.LANCZOS)

            # Convert to PhotoImage for display
            tk_image = ImageTk.PhotoImage(preview_image)

            # Update the preview label
            self.preview_label.config(image=tk_image)
            self.preview_label.image = tk_image  # Keep a reference to prevent garbage collection

        except Exception as e:
            self.status_var.set(f"Error loading image: {e}")
            messagebox.showerror("Error", f"Failed to load image: {e}")
    
    def print_image(self):
        if not self.image_path:
            return
        
        try:
            # Get settings from UI as integers
            threshold = int(self.threshold_var.get())
            position_x = int(self.pos_x_var.get())
            position_y = int(self.pos_y_var.get())
            print_speed = int(self.speed_var.get())
            darkness = int(self.darkness_var.get())
            quantity = int(self.quantity_var.get())
            pw, ll = self.label_sizes[self.label_size_var.get()]
            self.printer_name = self.printer_var.get()
            
            # Open and convert image
            img = self.pil_image.copy()
            
            # Convert to black and white
            if img.mode != '1':
                # Convert to grayscale first if needed
                if img.mode != 'L':
                    img = img.convert('L')
                
                # Convert to 1-bit with threshold (inverted for ZPL)
                img = img.point(lambda x: 0 if x >= threshold else 255, mode='1')
            
            # Get dimensions
            width, height = img.size
            
            # Calculate bytes per row (rounded up to nearest byte)
            bytes_per_row = (width + 7) // 8
            
            # Get the binary data
            binary_data = img.tobytes()
            
            # Convert binary data to hex
            hex_rows = []
            for i in range(0, len(binary_data), bytes_per_row):
                row = binary_data[i:i+bytes_per_row]
                hex_row = ''.join(f'{byte:02X}' for byte in row)
                hex_rows.append(hex_row)
            
            # Join all rows with newlines
            hex_data = '\n'.join(hex_rows)
            
            # Calculate total bytes
            total_bytes = bytes_per_row * height
            
            # Create ZPL code with quality settings - ensure all values are integers
            zpl = f"^XA\n"
            zpl += f"^PW{pw}\n^LL{ll}\n"           # Label width and length
            zpl += f"^PR{int(print_speed)},0,0\n"  # Print speed
            zpl += f"^MD{int(darkness)}\n"         # Darkness/Intensity
            zpl += f"^PQ{int(quantity)}\n"         # Quantity
            zpl += f"^FO{int(position_x)},{int(position_y)}\n"
            zpl += f"^GFA,{total_bytes},{total_bytes},{bytes_per_row},\n{hex_data}\n"
            zpl += "^FS\n^XZ"
            
            # Create a temporary file
            temp_file = "temp_print.zpl"
            with open(temp_file, 'w') as f:
                f.write(zpl)
            
            # Print using lp command
            cmd = ["lp", "-d", self.printer_name, "-o", "raw", temp_file]
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                self.status_var.set(f"Print job sent successfully")
                messagebox.showinfo("Print Job", f"Print job sent successfully to {self.printer_name}")
                
                # Show print information
                info_message = f"Image printed successfully!\n\n" \
                              f"Image size: {width}x{height}\n" \
                              f"Bytes per row: {bytes_per_row}\n" \
                              f"Total bytes: {total_bytes}\n" \
                              f"Print speed: {print_speed}\n" \
                              f"Darkness: {darkness}\n" \
                              f"Quantity: {quantity}"
                
                messagebox.showinfo("Print Details", info_message)
            else:
                error_msg = f"Error sending print job: {result.stderr}"
                self.status_var.set(error_msg)
                messagebox.showerror("Print Error", error_msg)
            
            # Clean up
            try:
                os.remove(temp_file)
            except:
                pass
        
        except Exception as e:
            error_msg = f"Error: {e}"
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)

def main():
    root = tk.Tk()
    app = ZPLPrinterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()