#!/usr/bin/env python3
"""
Template Capture Utility for VBS Automation
Provides tools for capturing and managing UI element templates
"""

import cv2
import numpy as np
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
import pyautogui
import time
import json
import os
from typing import Tuple, Optional, Dict, Any
from .template_service import TemplateService

class TemplateCaptureGUI:
    """GUI for capturing templates from screen"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VBS Template Capture Tool")
        self.root.geometry("400x300")
        
        self.template_service = TemplateService()
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the GUI interface"""
        # Title
        title_label = tk.Label(self.root, text="VBS Template Capture Tool", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Instructions
        instructions = tk.Label(self.root, 
                               text="1. Click 'Capture Template' to start\n"
                                    "2. Select area on screen to capture\n"
                                    "3. Enter template name and settings\n"
                                    "4. Template will be saved automatically",
                               justify=tk.LEFT)
        instructions.pack(pady=10)
        
        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20)
        
        capture_btn = tk.Button(button_frame, text="Capture Template", 
                               command=self.capture_template, 
                               bg="#4CAF50", fg="white", font=("Arial", 12))
        capture_btn.pack(pady=5)
        
        list_btn = tk.Button(button_frame, text="List Templates", 
                            command=self.list_templates,
                            bg="#2196F3", fg="white", font=("Arial", 12))
        list_btn.pack(pady=5)
        
        test_btn = tk.Button(button_frame, text="Test Template", 
                            command=self.test_template,
                            bg="#FF9800", fg="white", font=("Arial", 12))
        test_btn.pack(pady=5)
        
        quit_btn = tk.Button(button_frame, text="Quit", 
                            command=self.root.quit,
                            bg="#f44336", fg="white", font=("Arial", 12))
        quit_btn.pack(pady=5)
        
        # Status
        self.status_label = tk.Label(self.root, text="Ready to capture templates", 
                                    fg="green")
        self.status_label.pack(side=tk.BOTTOM, pady=10)
    
    def capture_template(self):
        """Capture template from screen selection"""
        try:
            self.status_label.config(text="Minimizing window in 3 seconds...", fg="orange")
            self.root.update()
            time.sleep(3)
            
            # Minimize window
            self.root.iconify()
            time.sleep(1)
            
            # Get screen selection
            selection = self.get_screen_selection()
            
            # Restore window
            self.root.deiconify()
            
            if selection:
                x, y, width, height = selection
                
                # Capture the selected area
                screenshot = pyautogui.screenshot(region=(x, y, width, height))
                template_image = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
                # Get template name
                template_name = simpledialog.askstring("Template Name", 
                                                      "Enter template name:")
                
                if template_name:
                    # Get additional settings
                    confidence = simpledialog.askfloat("Confidence Threshold", 
                                                      "Enter confidence threshold (0.0-1.0):",
                                                      initialvalue=0.8, minvalue=0.0, maxvalue=1.0)
                    
                    description = simpledialog.askstring("Description", 
                                                        "Enter template description (optional):",
                                                        initialvalue=f"Template for {template_name}")
                    
                    # Create metadata
                    metadata = {
                        'created_date': time.time(),
                        'usage_count': 0,
                        'success_rate': 0.0,
                        'confidence_threshold': confidence or 0.8,
                        'scale_factors': [1.0],
                        'description': description or f"Template for {template_name}",
                        'capture_region': {'x': x, 'y': y, 'width': width, 'height': height}
                    }
                    
                    # Save template
                    if self.template_service.save_template(template_name, template_image, metadata):
                        self.status_label.config(text=f"Template '{template_name}' saved successfully!", 
                                               fg="green")
                        messagebox.showinfo("Success", f"Template '{template_name}' saved successfully!")
                    else:
                        self.status_label.config(text="Failed to save template", fg="red")
                        messagebox.showerror("Error", "Failed to save template")
                else:
                    self.status_label.config(text="Template capture cancelled", fg="orange")
            else:
                self.status_label.config(text="No selection made", fg="orange")
                
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="red")
            messagebox.showerror("Error", f"Template capture failed: {str(e)}")
    
    def get_screen_selection(self) -> Optional[Tuple[int, int, int, int]]:
        """Get screen selection using mouse drag"""
        try:
            messagebox.showinfo("Screen Selection", 
                               "Click and drag to select the template area.\n"
                               "Press ESC to cancel.")
            
            # Create overlay window for selection
            overlay = tk.Toplevel()
            overlay.attributes('-fullscreen', True)
            overlay.attributes('-alpha', 0.3)
            overlay.configure(bg='black')
            overlay.attributes('-topmost', True)
            
            canvas = tk.Canvas(overlay, highlightthickness=0)
            canvas.pack(fill=tk.BOTH, expand=True)
            
            selection_rect = None
            start_x = start_y = 0
            
            def on_mouse_down(event):
                nonlocal start_x, start_y
                start_x, start_y = event.x_root, event.y_root
            
            def on_mouse_drag(event):
                nonlocal selection_rect
                if selection_rect:
                    canvas.delete(selection_rect)
                
                x1, y1 = start_x, start_y
                x2, y2 = event.x_root, event.y_root
                
                # Convert to canvas coordinates
                canvas_x1 = min(x1, x2)
                canvas_y1 = min(y1, y2)
                canvas_x2 = max(x1, x2)
                canvas_y2 = max(y1, y2)
                
                selection_rect = canvas.create_rectangle(
                    canvas_x1, canvas_y1, canvas_x2, canvas_y2,
                    outline='red', width=2, fill='red', stipple='gray50'
                )
            
            def on_mouse_up(event):
                overlay.quit()
            
            def on_key_press(event):
                if event.keysym == 'Escape':
                    overlay.quit()
            
            # Bind events
            overlay.bind('<Button-1>', on_mouse_down)
            overlay.bind('<B1-Motion>', on_mouse_drag)
            overlay.bind('<ButtonRelease-1>', on_mouse_up)
            overlay.bind('<Key>', on_key_press)
            overlay.focus_set()
            
            # Start selection
            overlay.mainloop()
            
            # Get final selection
            if selection_rect:
                coords = canvas.coords(selection_rect)
                if len(coords) == 4:
                    x1, y1, x2, y2 = coords
                    width = int(x2 - x1)
                    height = int(y2 - y1)
                    
                    overlay.destroy()
                    
                    if width > 10 and height > 10:  # Minimum size check
                        return (int(x1), int(y1), width, height)
            
            overlay.destroy()
            return None
            
        except Exception as e:
            print(f"Selection error: {str(e)}")
            return None
    
    def list_templates(self):
        """Show list of available templates"""
        templates = self.template_service.get_template_list()
        
        if not templates:
            messagebox.showinfo("Templates", "No templates found")
            return
        
        # Create template list window
        list_window = tk.Toplevel(self.root)
        list_window.title("Available Templates")
        list_window.geometry("500x400")
        
        # Create listbox with scrollbar
        frame = tk.Frame(list_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Add templates to list
        for template in sorted(templates):
            info = self.template_service.get_template_info(template)
            if info:
                usage = info.get('usage_count', 0)
                success_rate = info.get('success_rate', 0.0)
                listbox.insert(tk.END, f"{template} (Used: {usage}, Success: {success_rate:.1%})")
            else:
                listbox.insert(tk.END, template)
        
        # Add info button
        def show_template_info():
            selection = listbox.curselection()
            if selection:
                template_name = templates[selection[0]]
                info = self.template_service.get_template_info(template_name)
                if info:
                    info_text = f"Template: {template_name}\n\n"
                    info_text += f"Description: {info.get('description', 'N/A')}\n"
                    info_text += f"Created: {time.ctime(info.get('created_date', 0))}\n"
                    info_text += f"Usage Count: {info.get('usage_count', 0)}\n"
                    info_text += f"Success Rate: {info.get('success_rate', 0.0):.1%}\n"
                    info_text += f"Confidence Threshold: {info.get('confidence_threshold', 0.8)}\n"
                    
                    messagebox.showinfo(f"Template Info: {template_name}", info_text)
        
        info_btn = tk.Button(list_window, text="Show Info", command=show_template_info)
        info_btn.pack(pady=5)
    
    def test_template(self):
        """Test template matching"""
        templates = self.template_service.get_template_list()
        
        if not templates:
            messagebox.showinfo("Test Template", "No templates found")
            return
        
        # Create template selection window
        test_window = tk.Toplevel(self.root)
        test_window.title("Test Template")
        test_window.geometry("300x200")
        
        tk.Label(test_window, text="Select template to test:").pack(pady=10)
        
        # Template selection
        template_var = tk.StringVar(value=templates[0])
        template_menu = tk.OptionMenu(test_window, template_var, *templates)
        template_menu.pack(pady=5)
        
        # Test button
        def run_test():
            template_name = template_var.get()
            
            test_window.destroy()
            self.status_label.config(text="Testing template...", fg="orange")
            self.root.update()
            
            # Run template matching
            result = self.template_service.match_template(template_name)
            
            if result.success:
                match = result.matches[0]
                message = f"Template found!\n\n"
                message += f"Confidence: {match.confidence:.2f}\n"
                message += f"Location: {match.location}\n"
                message += f"Center: {match.center}\n"
                message += f"Processing time: {result.processing_time:.3f}s"
                
                self.status_label.config(text="Template test successful!", fg="green")
                messagebox.showinfo("Test Result", message)
            else:
                message = f"Template not found\n\n"
                message += f"Processing time: {result.processing_time:.3f}s"
                if result.error_message:
                    message += f"\nError: {result.error_message}"
                
                self.status_label.config(text="Template test failed", fg="red")
                messagebox.showwarning("Test Result", message)
        
        tk.Button(test_window, text="Run Test", command=run_test).pack(pady=20)
    
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()

def main():
    """Main entry point for template capture utility"""
    try:
        app = TemplateCaptureGUI()
        app.run()
    except Exception as e:
        print(f"Template capture utility error: {str(e)}")

if __name__ == "__main__":
    main()