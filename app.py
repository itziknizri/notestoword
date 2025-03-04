import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from docx import Document
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import tempfile
from datetime import datetime

class WordCommentsExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("ממיר הערות מוורד לאקסל")
        self.root.geometry("600x500")
        
        # ערכים שנשמור
        self.docx_path = None
        self.comments_data = []
        
        # יצירת ממשק
        self.create_widgets()
    
    def create_widgets(self):
        # כותרת
        header = tk.Label(self.root, text="ממיר הערות מקובץ וורד לאקסל", font=("Arial", 16, "bold"))
        header.pack(pady=20)
        
        # מסגרת לבחירת קובץ
        file_frame = tk.Frame(self.root)
        file_frame.pack(fill="x", padx=20, pady=10)
        
        self.file_label = tk.Label(file_frame, text="לא נבחר קובץ", width=40, anchor="w")
        self.file_label.pack(side="left", padx=5)
        
        select_btn = tk.Button(file_frame, text="בחר קובץ וורד", command=self.select_file)
        select_btn.pack(side="right", padx=5)
        
        # כפתור עיבוד
        process_btn = tk.Button(self.root, text="עבד את הקובץ", command=self.process_file, height=2)
        process_btn.pack(pady=20)
        
        # אזור סטטוס
        self.status_var = tk.StringVar()
        self.status_var.set("מוכן")
        status_label = tk.Label(self.root, textvariable=self.status_var, fg="blue")
        status_label.pack(pady=5)
        
        # מסגרת לתצוגת תוצאות
        results_frame = tk.Frame(self.root)
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # טבלת תוצאות
        self.result_tree = ttk.Treeview(results_frame, columns=("הערה", "כותב", "עמוד"), show="headings")
        
        # הגדרת כותרות
        self.result_tree.heading("הערה", text="הערה")
        self.result_tree.heading("כותב", text="כותב")
        self.result_tree.heading("עמוד", text="עמוד")
        
        # הגדרת רוחב עמודות
        self.result_tree.column("הערה", width=300)
        self.result_tree.column("כותב", width=100)
        self.result_tree.column("עמוד", width=50)
        
        # גלילה
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        # מיקום
        self.result_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # כפתור ייצוא
        export_btn = tk.Button(self.root, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(pady=10)
        
        # מידע
        info_label = tk.Label(self.root, text="פותח עם ❤️ לטובת ייצוא הערות מקובצי וורד", fg="gray")
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
                self.result_tree.insert("", "end", values=(
                    comment.get('הערה', '')[:50] + '...' if len(comment.get('הערה', '')) > 50 else comment.get('הערה', ''),
                    comment.get('כותב', ''),
                    comment.get('עמוד', '')
                ))
            
            self.comments_data = comments
            
            if comments:
                self.status_var.set(f"נמצאו {len(comments)} הערות. ניתן לייצא לאקסל.")
            else:
                self.status_var.set("לא נמצאו הערות בקובץ.")
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בעיבוד הקובץ:\n{str(e)}")
            self.status_var.set("אירעה שגיאה")
    
    def extract_comments_from_docx(self, docx_path):
        """
        מחלץ הערות ותגובות מקובץ וורד
        """
        # מידע שנרצה לשמור לגבי כל הערה
        comments_data = []
        
        try:
            # קובץ וורד הוא למעשה קובץ ZIP שמכיל מסמכי XML
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # בדיקה אם יש קובץ הערות
                comment_files = [f for f in zip_ref.namelist() if 'comments.xml' in f]
                
                if not comment_files:
                    return comments_data
                
                # קריאת מסמך הוורד לצורך קבלת הטקסט המקורי
                doc = Document(docx_path)
                paragraphs = [p.text for p in doc.paragraphs]
                
                # מיפוי מספרי פסקאות לעמודים (הערכה פשוטה)
                para_to_page = {}
                approx_chars_per_page = 3000  # הערכה של כמות תווים בעמוד
                current_page = 1
                char_count = 0
                
                for i, para in enumerate(paragraphs):
                    char_count += len(para)
                    para_to_page[i] = current_page
                    if char_count > approx_chars_per_page:
                        current_page += 1
                        char_count = 0
                
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
                        
                        # טקסט ההערה
                        comment_text_elements = comment.findall('.//w:t', ns)
                        comment_text = "".join([elem.text for elem in comment_text_elements if elem.text])
                        
                        # חיפוש תגובות (הערות מקושרות)
                        parent_id = comment.get(f"{{{ns['w']}}}parentId")
                        
                        # הוספת מידע ההערה לרשימה
                        comment_data = {
                            'id': comment_id,
                            'parent_id': parent_id,
                            'author': author,
                            'date': date,
                            'text': comment_text,
                            'page': 0  # ערך ברירת מחדל, יעודכן בהמשך
                        }
                        
                        comments_data.append(comment_data)
            
            # יצירת מבנה נתונים היררכי של הערות ותגובות
            comments_dict = {}
            for comment in comments_data:
                comments_dict[comment['id']] = comment
            
            # מסדר הערות ותגובות
            structured_comments = []
            
            # מיון ההערות לפי עיקריות ותגובות
            for comment in comments_data:
                if comment['parent_id'] is None:  # הערה עיקרית
                    # חיפוש כל התגובות להערה זו
                    replies = []
                    for reply in comments_data:
                        if reply['parent_id'] == comment['id']:
                            replies.append(reply)
                    
                    # יצירת שורה חדשה לאקסל
                    row = {
                        'הערה': comment['text'],
                        'כותב': comment['author'],
                        'עמוד': para_to_page.get(0, 1),  # כברירת מחדל עמוד 1
                        'תאריך': comment['date']
                    }
                    
                    # הוספת תגובות
                    for i, reply in enumerate(replies, 1):
                        row[f'תגובה {i}'] = reply['text']
                        row[f'כותב תגובה {i}'] = reply['author']
                        row[f'תאריך תגובה {i}'] = reply['date']
                    
                    structured_comments.append(row)
            
            return structured_comments
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בחילוץ ההערות:\n{str(e)}")
            return []
    
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
            
            # פתיחת התיקייה
            os.startfile(os.path.dirname(excel_path))
            
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בייצוא לאקסל:\n{str(e)}")

def main():
    root = tk.Tk()
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
