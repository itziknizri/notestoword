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
        
        # הגדרת התמיכה ב-RTL עבור עברית - שיפור משמעותי
        self.configure_rtl_support()
        
        # ערכים שנשמור
        self.docx_path = None
        self.comments_data = []
        
        # יצירת ממשק
        self.create_widgets()
        
        # שיפור המראה
        self.style_widgets()
    
    def configure_rtl_support(self):
        """הגדרת תמיכה מלאה ב-RTL"""
        # יצירת פונטים משופרים
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        
        # שינוי הכיוון ל-RTL
        self.root.tk.call('package', 'require', 'Ttk')
        
        try:
            # הגדרות RTL מתקדמות
            self.root.tk.call('tk', 'scaling', 1.0)
            # הפיכת כיוון הטקסט לימין-לשמאל
            self.root.tk.call('tk', 'scaling', 1.0)
            self.root.tk.call('option', 'add', '*TCombobox*Listbox.font', default_font)
            self.root.tk.call('option', 'add', '*TCombobox*Listbox.justify', 'right')
            
            # הגדרת כיוון למחלקות ספציפיות
            ttk_style = ttk.Style()
            ttk_style.configure('TNotebook.Tab', justify='right')
            ttk_style.configure('TButton', justify='right')
            ttk_style.configure('TLabel', justify='right')
            ttk_style.configure('Treeview', justify='right')
            ttk_style.configure('Treeview.Heading', justify='right')
        except tk.TclError:
            pass
    
    def style_widgets(self):
        # הגדרת סגנון לתוכנה
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Treeview", font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        
        # הגדרת כיוון טקסט לעברית
        style.configure("TButton", justify="right")
        style.configure("TLabel", justify="right")
        
        # הגדרת צבעים
        self.root.configure(bg="#f0f0f0")
    
    def create_widgets(self):
        # מסגרת ראשית
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # כותרת
        header = ttk.Label(main_frame, text="ממיר הערות מקובץ וורד לאקסל", style="Header.TLabel", anchor="center")
        header.pack(pady=20)
        
        # מסגרת לבחירת קובץ
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=10)
        
        # שינוי סדר הרכיבים לתמיכה בRTL
        select_btn = ttk.Button(file_frame, text="בחר קובץ וורד", command=self.select_file)
        select_btn.pack(side="right", padx=5)
        
        self.file_label = ttk.Label(file_frame, text="לא נבחר קובץ", width=60, anchor="e")
        self.file_label.pack(side="left", padx=5, fill="x", expand=True)
        
        # כפתור עיבוד
        process_btn = ttk.Button(main_frame, text="עבד את הקובץ", command=self.process_file)
        process_btn.pack(pady=10)
        
        # אזור סטטוס עם מסגרת
        status_frame = ttk.Frame(main_frame, relief="sunken", borderwidth=1)
        status_frame.pack(fill="x", pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("מוכן")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="e")
        status_label.pack(pady=5, fill="x")
        
        # מסגרת לתצוגת תוצאות
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(fill="both", expand=True, pady=10)
        
        # טבלת תוצאות - תמיכה בשרשור נכון ואינדקס
        columns = ("#", "הערה", "כותב", "עמוד", "תאריך", "תגובה 1", "כותב תגובה 1", "תאריך 1", 
                  "תגובה 2", "כותב תגובה 2", "תאריך 2", "תגובה 3", "כותב תגובה 3", "תאריך 3")
        self.result_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        # הגדרת כותרות
        for col in columns:
            self.result_tree.heading(col, text=col)
        
        # הגדרת רוחב עמודות
        self.result_tree.column("#", width=40, anchor="e")
        self.result_tree.column("הערה", width=200, anchor="e")
        self.result_tree.column("כותב", width=80, anchor="e")
        self.result_tree.column("עמוד", width=50, anchor="e")
        self.result_tree.column("תאריך", width=120, anchor="e")
        self.result_tree.column("תגובה 1", width=200, anchor="e")
        self.result_tree.column("כותב תגובה 1", width=80, anchor="e")
        self.result_tree.column("תאריך 1", width=120, anchor="e")
        self.result_tree.column("תגובה 2", width=200, anchor="e")
        self.result_tree.column("כותב תגובה 2", width=80, anchor="e")
        self.result_tree.column("תאריך 2", width=120, anchor="e")
        self.result_tree.column("תגובה 3", width=200, anchor="e")
        self.result_tree.column("כותב תגובה 3", width=80, anchor="e")
        self.result_tree.column("תאריך 3", width=120, anchor="e")
        
        # גלילה אופקית ואנכית
        x_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=self.result_tree.xview)
        y_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # מיקום (שימוש ב-pack במקום grid לתמיכה טובה יותר בRTL)
        y_scrollbar.pack(side="right", fill="y")
        x_scrollbar.pack(side="bottom", fill="x")
        self.result_tree.pack(side="left", fill="both", expand=True)
        
        # כפתורי ייצוא ויציאה
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        # שימוש בLTR לכפתורים בשורה התחתונה
        export_btn = ttk.Button(button_frame, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(side="right", padx=5)
        
        exit_btn = ttk.Button(button_frame, text="יציאה", command=self.root.quit)
        exit_btn.pack(side="left", padx=5)
        
        # מידע - כולל קרדיט מבוקש
        info_label = ttk.Label(main_frame, text="פותח ע\"י יצחק נזרי", foreground="gray", anchor="center")
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
            for idx, comment in enumerate(comments, 1):
                values = [
                    idx,  # אינדקס מספרי להערה
                    self.truncate_text(comment.get('הערה', ''), 50),
                    comment.get('כותב', ''),
                    comment.get('עמוד', ''),
                    comment.get('תאריך', ''),
                    self.truncate_text(comment.get('תגובה 1', ''), 50),
                    comment.get('כותב תגובה 1', ''),
                    comment.get('תאריך תגובה 1', ''),
                    self.truncate_text(comment.get('תגובה 2', ''), 50),
                    comment.get('כותב תגובה 2', ''),
                    comment.get('תאריך תגובה 2', ''),
                    self.truncate_text(comment.get('תגובה 3', ''), 50),
                    comment.get('כותב תגובה 3', ''),
                    comment.get('תאריך תגובה 3', '')
                ]
                self.result_tree.insert("", "end", values=values)
                
                # שמירת האינדקס בנתונים
                comment['אינדקס'] = idx
            
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
        מחלץ הערות ותגובות מקובץ וורד בצורה משופרת
        """
        # מידע שנרצה לשמור לגבי כל הערה
        comments_data = []
        
        try:
            # קריאת המסמך באמצעות python-docx
            doc = Document(docx_path)
            
            # אלגוריתם משופר יותר לחישוב מספרי עמודים
            # נחשב את ממוצע מספר התווים בעמוד (הערכה טובה יותר)
            total_chars = sum(len(p.text) for p in doc.paragraphs)
            
            # אומדן מספר העמודים באמצעות גישה מתקדמת יותר
            # ממוצע של כ-1800 תווים לעמוד (מקובל לתכנון עמודים)
            chars_per_page_estimate = 1800
            estimated_total_pages = max(1, total_chars // chars_per_page_estimate)
            
            # יצירת מיפוי של זמן יחסי בטקסט למספר העמוד
            page_mapping = {}
            
            # קובץ וורד הוא למעשה קובץ ZIP שמכיל מסמכי XML
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # בדיקה אם יש קובץ הערות
                comment_files = [f for f in zip_ref.namelist() if 'comments.xml' in f]
                
                if not comment_files:
                    return []
                
                # איסוף מידע על מיקום מותאם הערות (במידת האפשר)
                comment_locations = self.get_comment_locations(doc)
                
                # שמירת מידע על כל ההערות והתגובות
                all_comments = []
                comment_threads = {}  # מילון לשמירת שרשורי הערות
                
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
                        comment_text = "".join([elem.text if elem.text else "" for elem in comment_text_elements])
                        
                        # חיפוש תגובות (הערות מקושרות)
                        parent_id = comment.get(f"{{{ns['w']}}}parentId")
                        
                        # חישוב מספר העמוד
                        comment_pos = comment_locations.get(comment_id, 0)
                        relative_pos = min(1.0, max(0.0, comment_pos / total_chars if total_chars > 0 else 0))
                        page = max(1, min(estimated_total_pages, int(relative_pos * estimated_total_pages) + 1))
                        
                        # הוספת מידע ההערה לרשימה
                        comment_info = {
                            'id': comment_id,
                            'parent_id': parent_id,
                            'author': author,
                            'date': date_formatted,
                            'text': comment_text,
                            'page': page
                        }
                        
                        all_comments.append(comment_info)
                        
                        # ארגון ההערות בשרשורים
                        if parent_id is None:
                            # זו הערה ראשית
                            if comment_id not in comment_threads:
                                comment_threads[comment_id] = {
                                    'main': comment_info,
                                    'replies': []
                                }
                        else:
                            # זו תגובה
                            if parent_id not in comment_threads:
                                comment_threads[parent_id] = {
                                    'main': None,
                                    'replies': [comment_info]
                                }
                            else:
                                comment_threads[parent_id]['replies'].append(comment_info)
                
                # ארגון ההערות בפורמט הסופי לתצוגה וייצוא
                result_comments = []
                
                for thread_id, thread in comment_threads.items():
                    if thread['main'] is None:
                        # במקרה נדיר שאין הערה ראשית
                        continue
                    
                    # מיון התגובות לפי תאריך
                    replies = sorted(thread['replies'], key=lambda x: x.get('date', ''))
                    
                    # יצירת שורה חדשה לאקסל
                    row = {
                        'הערה': thread['main']['text'],
                        'כותב': thread['main']['author'],
                        'עמוד': thread['main']['page'],
                        'תאריך': thread['main']['date']
                    }
                    
                    # הוספת תגובות כעמודות
                    for i, reply in enumerate(replies, 1):
                        if i <= 10:  # מוגבל ל-10 תגובות לכל היותר
                            row[f'תגובה {i}'] = reply['text']
                            row[f'כותב תגובה {i}'] = reply['author']
                            row[f'תאריך תגובה {i}'] = reply['date']
                    
                    result_comments.append(row)
                
                # מיון לפי מספר עמוד
                result_comments.sort(key=lambda x: (x.get('עמוד', 0), x.get('תאריך', '')))
                
                return result_comments
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בחילוץ ההערות:\n{str(e)}")
            return []
    
    def get_comment_locations(self, doc):
        """
        מנסה לאתר את המיקום של כל הערה בטקסט
        """
        comment_locations = {}
        char_position = 0
        
        # נעבור על כל הפסקאות
        for paragraph in doc.paragraphs:
            # נוסיף את אורך הטקסט הנוכחי
            paragraph_text_length = len(paragraph.text)
            
            # נבדוק אם יש הערות בפסקה
            for run in paragraph.runs:
                if hasattr(run, '_element') and run._element.xpath('.//w:commentReference'):
                    for comment_ref in run._element.xpath('.//w:commentReference'):
                        comment_id = comment_ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                        if comment_id:
                            # שמירת המיקום המוערך של ההערה
                            comment_locations[comment_id] = char_position
            
            char_position += paragraph_text_length
        
        return comment_locations
    
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
            
            # סידור העמודות בסדר הנכון
            columns_order = ['אינדקס', 'הערה', 'כותב', 'עמוד', 'תאריך']
            
            # הוספת עמודות תגובה
            for i in range(1, 11):  # תמיכה בעד 10 תגובות
                reply_cols = [f'תגובה {i}', f'כותב תגובה {i}', f'תאריך תגובה {i}']
                columns_order.extend([col for col in reply_cols if col in df.columns])
            
            # סינון עמודות קיימות בדאטה
            final_columns = [col for col in columns_order if col in df.columns]
            
            # שמירה לאקסל עם סדר העמודות הנכון
            df = df[final_columns]
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
    # הגדרת RTL כברירת מחדל
    root.tk.call('tk', 'windowingsystem')  # נדרש לפני שימוש ב-tk_strictMotif
    # הגדרות לתמיכה מלאה בRTL
    try:
        root.tk_strictMotif(False)
    except:
        pass
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
