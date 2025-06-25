import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from tkinter import font

class PyWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PyWord - MS Word-like App")
        self.text = tk.Text(root, wrap='word', undo=True)
        self.text.pack(expand=1, fill='both')
        self.create_menu()
        self.create_toolbar()
        self.create_statusbar()
        self.text.bind('<KeyRelease>', self.update_statusbar)
        self.text.bind('<ButtonRelease>', self.update_statusbar)
        self.bold_font = font.Font(self.text, self.text.cget("font"))
        self.bold_font.configure(weight="bold")
        self.italic_font = font.Font(self.text, self.text.cget("font"))
        self.italic_font.configure(slant="italic")
        self.underline_font = font.Font(self.text, self.text.cget("font"))
        self.underline_font.configure(underline=1)
        self.text.tag_configure("bold", font=self.bold_font)
        self.text.tag_configure("italic", font=self.italic_font)
        self.text.tag_configure("underline", font=self.underline_font)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open", command=self.open_docx)
        filemenu.add_command(label="Save As", command=self.save_docx)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        self.root.config(menu=menubar)

    def create_toolbar(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        bold_btn = tk.Button(toolbar, text="Bold", command=self.make_bold)
        bold_btn.pack(side=tk.LEFT, padx=2, pady=2)
        italic_btn = tk.Button(toolbar, text="Italic", command=self.make_italic)
        italic_btn.pack(side=tk.LEFT, padx=2, pady=2)
        underline_btn = tk.Button(toolbar, text="Underline", command=self.make_underline)
        underline_btn.pack(side=tk.LEFT, padx=2, pady=2)
        toolbar.pack(side=tk.TOP, fill=tk.X)

    def create_statusbar(self):
        self.statusbar = tk.Label(self.root, text="Ln 1, Col 1", anchor='w')
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

    def update_statusbar(self, event=None):
        row, col = self.text.index(tk.INSERT).split('.')
        self.statusbar.config(text=f"Ln {int(row)}, Col {int(col)+1}")

    def make_bold(self):
        self.toggle_tag("bold")

    def make_italic(self):
        self.toggle_tag("italic")

    def make_underline(self):
        self.toggle_tag("underline")

    def toggle_tag(self, tag):
        try:
            start, end = self.text.index(tk.SEL_FIRST), self.text.index(tk.SEL_LAST)
            if tag in self.text.tag_names(tk.SEL_FIRST):
                self.text.tag_remove(tag, start, end)
            else:
                self.text.tag_add(tag, start, end)
        except tk.TclError:
            pass  # No selection

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
