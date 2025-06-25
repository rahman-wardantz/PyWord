import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document

class PyWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PyWord - MS Word-like App")
        self.text = tk.Text(root, wrap='word', undo=True)
        self.text.pack(expand=1, fill='both')
        self.create_menu()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.open_docx)
        filemenu.add_command(label="Save As", command=self.save_docx)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        self.root.config(menu=menubar)

    def open_docx(self):
        filepath = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if filepath:
            try:
                doc = Document(filepath)
                self.text.delete(1.0, tk.END)
                for para in doc.paragraphs:
                    self.text.insert(tk.END, para.text + '\n')
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {e}")

    def save_docx(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if filepath:
            try:
                doc = Document()
                content = self.text.get(1.0, tk.END).strip().split('\n')
                for line in content:
                    doc.add_paragraph(line)
                doc.save(filepath)
                messagebox.showinfo("Saved", "File saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

def main():
    root = tk.Tk()
    app = PyWordApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
