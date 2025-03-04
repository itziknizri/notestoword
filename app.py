import tkinter as tk
from tkinter import filedialog, ttk, messagebox, font
from docx import Document
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import tempfile
from datetime import datetime
import docx

class WordCommentsExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("ממיר הערות מוורד לאקסל")
        self.root.geometry("800x600")
        
        # הגדרת התמיכה ב-RTL עבור עברית
        self.configure_rtl_support()
        
        # ערכים שנשמור
        self.docx_path = None
        self.comments_data = []
        
        # יצירת ממשק
        self.create_widgets()
        
        # שיפור המראה
        self.style_widgets()
    
    def configure_rtl_support(self):
        # יצירת פונטים משופרים
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        
        # שינוי הכיוון ל-RTL
        self.root.tk.call('package', 'require', 'Ttk')
        
        try:
            # לא כל מערכות ההפעלה תומכות ב-RTL באופן מלא
            self.root.tk.call('tk', 'scaling', 1.0)
            self.root.tk.call('ttk::style', 'configure', '.', 'justify', 'right')
        except tk.TclError:
            pass
    
    def style_widgets(self):
        # הגדרת סגנון לתוכנה
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        
        # הגדרת צבעים
        self.root.configure(bg="#f0f0f0")
    
    def create_widgets(self):
        # מסגרת ראשית
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # כותרת
        header = ttk.Label(main_frame, text="ממיר הערות מקובץ וורד לאקסל", style="Header.TLabel")
        header.pack(pady=20)
        
        # מסגרת לבחירת קובץ
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=10)
        
        self.file_label = ttk.Label(file_frame, text="לא נבחר קובץ", width=60)
        self.file_label.pack(side="right", padx=5)
        
        select_btn = ttk.Button(file_frame, text="בחר קובץ וורד", command=self.select_file)
        select_btn.pack(side="left", padx=5)
        
        # כפתור עיבוד
        process_btn = ttk.Button(main_frame, text="עבד את הקובץ", command=self.process_file)
        process_btn.pack(pady=10)
        
        # אזור סטטוס עם מסגרת
        status_frame = ttk.Frame(main_frame, relief="sunken", borderwidth=1)
        status_frame.pack(fill="x", pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("מוכן")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(pady=5)
        
        # מסגרת לתצוגת תוצאות
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(fill="both", expand=True, pady=10)
        
        # טבלת תוצאות - נשנה כדי להציג עמודות רבות יותר
        columns = ("הערה", "כותב", "עמוד", "תאריך", "תגובה 1", "כותב תגובה 1", "תגובה 2", "כותב תגובה 2")
        self.result_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        # הגדרת כותרות
        for col in columns:
            self.result_tree.heading(col, text=col)
        
        # הגדרת רוחב עמודות
        self.result_tree.column("הערה", width=250)
        self.result_tree.column("כותב", width=80)
        self.result_tree.column("עמוד", width=50)
        self.result_tree.column("תאריך", width=120)
        self.result_tree.column("תגובה 1", width=200)
        self.result_tree.column("כותב תגובה 1", width=80)
        self.result_tree.column("תגובה 2", width=200)
        self.result_tree.column("כותב תגובה 2", width=80)
        
        # גלילה אופקית ואנכית
        x_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=self.result_tree.xview)
        y_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # מיקום
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        
        # הגדרת גדילה
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        # כפתור ייצוא
        export_frame = ttk.Frame(main_frame)
        export_frame.pack(fill="x", pady=10)
        
        export_btn = ttk.Button(export_frame, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(side="left", padx=5)
        
        # כפתור יציאה
        exit_btn = ttk.Button(export_frame, text="יציאה", command=self.root.quit)
        exit_btn.pack(side="right", padx=5)
        
        # מידע
        info_label = ttk.Label(main_frame, text="פותח עם ❤️ לטובת ייצוא הערות מקובצי וורד", foreground="gray")
        info_label.pack(pady=5)
    
    def select_file(self):
        self.docx_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")],
            title="בחר קובץ וורד"
        )
        
        if self.docx_path:
            self.file_label.config(text=os.path.basename(self.docx_path))
            self.status_var.set("קובץ נבחר. לחץ על 'עבד את הקובץ' להמשך.")
    
    def process_file(self):
        if not self.docx_path:
            messagebox.showwarning("שגיאה", "יש לבחור קובץ וורד תחילה")
            return
        
        self.status_var.set("מעבד את הקובץ...")
        self.root.update()
        
        try:
            # ניקוי נתונים קודמים
            self.result_tree.delete(*self.result_tree.get_children())
            self.comments_data = []
            
            # הפעלת הפונקציה לחילוץ הערות
            comments = self.extract_comments_from_docx(self.docx_path)
            
            # הצגת התוצאות בטבלה
            for idx, comment in enumerate(comments):
                values = [
                    self.truncate_text(comment.get('הערה', ''), 50),
                    comment.get('כותב', ''),
                    comment.get('עמוד', ''),
                    comment.get('תאריך', ''),
                    self.truncate_text(comment.get('תגובה 1', ''), 50),
                    comment.get('כותב תגובה 1', ''),
                    self.truncate_text(comment.get('תגובה 2', ''), 50),
                    comment.get('כותב תגובה 2', '')
                ]
                self.result_tree.insert("", "end", values=values)
            
            self.comments_data = comments
            
            if comments:
                self.status_var.set(f"נמצאו {len(comments)} הערות. ניתן לייצא לאקסל.")
            else:
                self.status_var.set("לא נמצאו הערות בקובץ.")
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בעיבוד הקובץ:\n{str(e)}")
            self.status_var.set("אירעה שגיאה")
    
    def truncate_text(self, text, length=50):
        """קיצור טקסט ארוך לתצוגה בטבלה"""
        if not text:
            return ""
        return text[:length] + '...' if len(text) > length else text
    
    def extract_comments_from_docx(self, docx_path):
        """
        מחלץ הערות ותגובות מקובץ וורד
        """
        # מידע שנרצה לשמור לגבי כל הערה
        comments_data = []
        
        try:
            # קריאת המסמך באמצעות python-docx
            doc = Document(docx_path)
            
            # אלגוריתם משופר לחישוב מספרי עמודים
            # נחשב את ממוצע מספר התווים בעמוד
            total_chars = sum(len(p.text) for p in doc.paragraphs)
            approx_pages = max(1, total_chars // 3000)  # הערכה ראשונית: 3000 תווים לעמוד
            chars_per_page = total_chars / approx_pages
            
            # יצירת מיפוי של פסקאות לעמודים
            para_to_page = {}
            current_page = 1
            char_count = 0
            
            for i, para in enumerate(doc.paragraphs):
                char_count += len(para.text)
                para_to_page[i] = current_page
                if char_count >= chars_per_page:
                    current_page += 1
                    char_count = 0
            
            # קובץ וורד הוא למעשה קובץ ZIP שמכיל מסמכי XML
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # בדיקה אם יש קובץ הערות
                comment_files = [f for f in zip_ref.namelist() if 'comments.xml' in f]
                
                if not comment_files:
                    return []
                
                # שמירת מידע על כל ההערות והתגובות
                all_comments = []
                
                # קריאת קובץ XML של ההערות
                for comment_file in comment_files:
                    xml_content = zip_ref.read(comment_file)
                    root = ET.fromstring(xml_content)
                    
                    # מציאת המרחב שמות (namespace)
                    ns = {'w': re.search(r'{(.*)}', root.tag).group(1)}
                    
                    # חילוץ כל ההערות
                    comments = root.findall('.//w:comment', ns)
                    
                    for comment in comments:
                        comment_id = comment.get(f"{{{ns['w']}}}id")
                        author = comment.get(f"{{{ns['w']}}}author", "לא ידוע")
                        date = comment.get(f"{{{ns['w']}}}date", "")
                        date_formatted = self.format_date(date)
                        
                        # טקסט ההערה
                        comment_text_elements = comment.findall('.//w:t', ns)
                        comment_text = "".join([elem.text for elem in comment_text_elements if elem.text])
                        
                        # חיפוש תגובות (הערות מקושרות)
                        parent_id = comment.get(f"{{{ns['w']}}}parentId")
                        
                        # מציאת מיקום ההערה (משוער)
                        # בהערכה פשוטה נשים את ההערה בעמוד העוקב אחרי העמוד הקודם לקבל פיזור
                        page = current_page // 2 if parent_id is None else 0
                        
                        # הוספת מידע ההערה לרשימה
                        all_comments.append({
                            'id': comment_id,
                            'parent_id': parent_id,
                            'author': author,
                            'date': date_formatted,
                            'text': comment_text,
                            'page': page
                        })
                
                # ארגון ההערות והתגובות בצורה טובה יותר
                # נשמור הערות עיקריות ונצרף אליהן את התגובות כעמודות
                parent_comments = []
                
                # מיפוי של הערות לפי ID
                comments_by_id = {c['id']: c for c in all_comments}
                
                # קבלת כל ההערות הראשיות (ללא parent_id)
                for comment in all_comments:
                    if comment['parent_id'] is None:
                        # יצירת שורה חדשה לאקסל
                        row = {
                            'הערה': comment['text'],
                            'כותב': comment['author'],
                            'עמוד': para_to_page.get(min(3, len(para_to_page)-1), 1),  # חישוב משופר לעמוד
                            'תאריך': comment['date']
                        }
                        
                        # חיפוש תגובות להערה זו
                        replies = [c for c in all_comments if c['parent_id'] == comment['id']]
                        
                        # הוספת תגובות כעמודות נוספות
                        for i, reply in enumerate(replies, 1):
                            row[f'תגובה {i}'] = reply['text']
                            row[f'כותב תגובה {i}'] = reply['author']
                            row[f'תאריך תגובה {i}'] = reply['date']
                        
                        parent_comments.append(row)
                
                return parent_comments
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בחילוץ ההערות:\n{str(e)}")
            return []
    
    def format_date(self, date_str):
        """פורמט תאריכים בצורה קריאה"""
        if not date_str:
            return ""
        
        try:
            # פורמט התאריך מגיע מ-Word בפורמט ISO
            date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            return date_obj.strftime('%d/%m/%Y %H:%M')
        except:
            return date_str
    
    def export_to_excel(self):
        if not self.comments_data:
            messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
            return
        
        # שאלה על שם הקובץ ומיקום
        excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="שמור קובץ אקסל",
            initialfile=f"comments_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not excel_path:
            return  # המשתמש ביטל
        
        try:
            # יצירת DataFrame
            df = pd.DataFrame(self.comments_data)
            
            # שמירה לאקסל
            df.to_excel(excel_path, index=False, engine='openpyxl')
            
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{excel_path}")
            self.status_var.set("הנתונים יוצאו בהצלחה")
            
            # נסיון לפתוח את התיקייה
            try:
                os.startfile(os.path.dirname(excel_path))
            except:
                pass
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בייצוא לאקסל:\n{str(e)}")

def main():
    root = tk.Tk()
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
