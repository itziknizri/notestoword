import tkinter as tk
from tkinter import filedialog, ttk, messagebox, font
from docx import Document
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
from datetime import datetime
import tempfile

class WordCommentsExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("ממיר הערות מוורד לאקסל")
        self.root.geometry("1100x650")
        
        # ערכים שנשמור
        self.docx_path = None
        self.comments_data = []
        
        # הגדרת כיוון RTL
        self.configure_rtl_support()
        
        # יצירת ממשק
        self.create_widgets()
    
    def configure_rtl_support(self):
        """הגדרת תמיכה בכיוון מימין לשמאל"""
        # יצירת פונטים
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        
        # הגדרת RTL באופן מערכתי
        try:
            self.root.tk_strictMotif(False)
        except:
            pass
        
        # הגדרת סגנון טקסט לימין
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10))
        style.configure("TLabel", anchor="e")  # טקסט לימין
        style.configure("TButton", anchor="e")  # טקסט לימין
        
        # הגדרת צבעים
        self.root.configure(bg="#f0f0f0")
    
    def create_widgets(self):
        # מסגרת ראשית
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # כותרת
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill="x", pady=10)
        
        header = ttk.Label(header_frame, text="ממיר הערות מקובץ וורד לאקסל", 
                          font=("Segoe UI", 16, "bold"))
        header.pack(side="right", padx=10)
        
        # מסגרת לבחירת קובץ (מסודרת מימין לשמאל)
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=10)
        
        select_btn = ttk.Button(file_frame, text="בחר קובץ וורד", command=self.select_file)
        select_btn.pack(side="right", padx=5)
        
        self.file_label = ttk.Label(file_frame, text="לא נבחר קובץ", anchor="e")
        self.file_label.pack(side="right", padx=5, fill="x", expand=True)
        
        # כפתור עיבוד
        process_frame = ttk.Frame(main_frame)
        process_frame.pack(fill="x", pady=5)
        
        process_btn = ttk.Button(process_frame, text="עבד את הקובץ", command=self.process_file)
        process_btn.pack(side="right", padx=5)
        
        # אזור סטטוס
        status_frame = ttk.Frame(main_frame, relief="sunken", borderwidth=1)
        status_frame.pack(fill="x", pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("מוכן")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="e")
        status_label.pack(side="right", pady=5, fill="x", expand=True)
        
        # יצירת מסגרת לטבלה
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True, pady=10)
        
        # הגדרת עמודות בסדר הנכון (מימין לשמאל)
        columns = (
            "תאריך תגובה 5", "כותב תגובה 5", "תגובה 5",
            "תאריך תגובה 4", "כותב תגובה 4", "תגובה 4",
            "תאריך תגובה 3", "כותב תגובה 3", "תגובה 3",
            "תאריך תגובה 2", "כותב תגובה 2", "תגובה 2",
            "תאריך תגובה 1", "כותב תגובה 1", "תגובה 1",
            "תאריך", "כותב", "הערה", "עמוד", "מס'"
        )
        
        # יצירת טבלה עם סדר עמודות מימין לשמאל
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        
        # הגדרת כותרות
        for col in columns:
            self.result_tree.heading(col, text=col, anchor="e")  # כותרות מיושרות לימין
            
            # רוחב קבוע לעמודות
            if col == "מס'":
                width = 40
            elif col == "עמוד":
                width = 50
            elif "תאריך" in col:
                width = 100
            elif "כותב" in col:
                width = 80
            elif "הערה" in col or "תגובה" in col:
                width = 200
            else:
                width = 100
                
            self.result_tree.column(col, width=width, anchor="e", stretch=False)  # רוחב קבוע ללא מתיחה
        
        # הגדרת סרגלי גלילה
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        y_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # סידור רכיבי הטבלה - RTL סדר
        y_scrollbar.pack(side="left", fill="y")  # סרגל גלילה אנכי בצד שמאל
        self.result_tree.pack(side="left", fill="both", expand=True)
        x_scrollbar.pack(side="bottom", fill="x")
        
        # מסגרת כפתורים
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        # כפתור ייצוא
        export_btn = ttk.Button(button_frame, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(side="right", padx=5)
        
        # כפתור יציאה
        exit_btn = ttk.Button(button_frame, text="יציאה", command=self.root.quit)
        exit_btn.pack(side="left", padx=5)
        
        # קרדיט
        credit_frame = ttk.Frame(main_frame)
        credit_frame.pack(fill="x", pady=5)
        
        credit = ttk.Label(credit_frame, text="פותח ע\"י יצחק נזרי", 
                          font=("Segoe UI", 9), foreground="gray")
        credit.pack(side="bottom", pady=5)
    
    def select_file(self):
        """בחירת קובץ וורד לעיבוד"""
        self.docx_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")],
            title="בחר קובץ וורד"
        )
        
        if self.docx_path:
            self.file_label.config(text=os.path.basename(self.docx_path))
            self.status_var.set("קובץ נבחר. לחץ על 'עבד את הקובץ' להמשך.")
    
    def process_file(self):
        """עיבוד הקובץ והצגת ההערות בטבלה"""
        if not self.docx_path:
            messagebox.showwarning("שגיאה", "יש לבחור קובץ וורד תחילה")
            return
        
        self.status_var.set("מעבד את הקובץ...")
        self.root.update()
        
        try:
            # ניקוי טבלה קודמת
            self.result_tree.delete(*self.result_tree.get_children())
            self.comments_data = []
            
            # חילוץ שרשורי התגובות - הערה ראשית עם כל התגובות שלה
            comment_threads = self.extract_comment_threads()
            
            # הצגת התוצאות בטבלה
            for idx, thread in enumerate(comment_threads, 1):
                # הכנת ערכים להצגה בטבלה - סדר מימין לשמאל
                values = [
                    thread.get('תאריך תגובה 5', ''),
                    thread.get('כותב תגובה 5', ''),
                    self.truncate_text(thread.get('תגובה 5', ''), this_max=50),
                    thread.get('תאריך תגובה 4', ''),
                    thread.get('כותב תגובה 4', ''),
                    self.truncate_text(thread.get('תגובה 4', ''), this_max=50),
                    thread.get('תאריך תגובה 3', ''),
                    thread.get('כותב תגובה 3', ''),
                    self.truncate_text(thread.get('תגובה 3', ''), this_max=50),
                    thread.get('תאריך תגובה 2', ''),
                    thread.get('כותב תגובה 2', ''),
                    self.truncate_text(thread.get('תגובה 2', ''), this_max=50),
                    thread.get('תאריך תגובה 1', ''),
                    thread.get('כותב תגובה 1', ''),
                    self.truncate_text(thread.get('תגובה 1', ''), this_max=50),
                    thread.get('תאריך', ''),
                    thread.get('כותב', ''),
                    self.truncate_text(thread.get('הערה', ''), this_max=50),
                    thread.get('עמוד', ''),
                    idx
                ]
                
                # הוספה לטבלה
                item_id = self.result_tree.insert("", 0, values=values)  # הוספה בתחילת הטבלה לסדר נכון
                
                # שמירת האינדקס למידע
                thread['מס\''] = idx
            
            self.comments_data = comment_threads
            
            if comment_threads:
                self.status_var.set(f"נמצאו {len(comment_threads)} שרשורי הערות. ניתן לייצא לאקסל.")
            else:
                self.status_var.set("לא נמצאו הערות בקובץ.")
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בעיבוד הקובץ:\n{str(e)}")
            self.status_var.set("אירעה שגיאה")
    
    def truncate_text(self, text, this_max=50):
        """קיצור טקסט ארוך לתצוגה בטבלה"""
        if not text:
            return ""
        return text[:this_max] + '...' if len(text) > this_max else text
    
    def extract_comment_threads(self):
        """
        מחלץ את כל ההערות והתגובות ומארגן אותן בשרשורים
        שרשור = הערה ראשית + כל התגובות שלה בעמודות נפרדות
        """
        if not self.docx_path or not os.path.exists(self.docx_path):
            return []
            
        try:
            all_comments = []  # כל ההערות והתגובות
            comment_threads = []  # שרשורי הערות לתצוגה
            page_map = {}  # מיפוי של הערות למספרי עמודים
            
            # פתיחת קובץ docx כארכיון ZIP
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                # חישוב עמודים מהמסמך
                try:
                    doc = Document(self.docx_path)
                    self.calculate_page_numbers(doc, page_map)
                except:
                    pass
                
                # חיפוש קבצי XML שמכילים הערות
                comment_files = [f for f in docx_zip.namelist() 
                               if 'word/comments' in f.lower() or 'word/comment' in f.lower()]
                
                if not comment_files:
                    return []
                
                # מיפוי של מזהי הערות להערות עצמן
                comment_map = {}
                
                # קריאת כל קבצי ההערות
                for comment_file in comment_files:
                    xml_content = docx_zip.read(comment_file)
                    root = ET.fromstring(xml_content)
                    
                    # זיהוי namespace
                    ns_match = re.search(r'{(.*?)}', root.tag)
                    if not ns_match:
                        continue
                        
                    ns = {'w': ns_match.group(1)}
                    
                    # חילוץ פרטי ההערות
                    for comment in root.findall('.//w:comment', ns):
                        comment_id = comment.get(f"{{{ns['w']}}}id")
                        if not comment_id:
                            continue
                            
                        # בדיקה אם זו תגובה להערה אחרת
                        parent_id = comment.get(f"{{{ns['w']}}}parentId")
                        
                        author = comment.get(f"{{{ns['w']}}}author", "")
                        date = comment.get(f"{{{ns['w']}}}date", "")
                        
                        # חילוץ טקסט ההערה
                        text_parts = []
                        for text_elem in comment.findall('.//w:t', ns):
                            if text_elem.text:
                                text_parts.append(text_elem.text)
                        
                        comment_text = "".join(text_parts)
                        
                        # חישוב מספר עמוד (אם לא חושב קודם)
                        page = page_map.get(comment_id, 1)
                        
                        # יצירת אובייקט ההערה
                        comment_obj = {
                            'id': comment_id,
                            'parent_id': parent_id,
                            'author': author,
                            'date': self.format_date(date),
                            'text': comment_text,
                            'page': page,
                            'is_reply': parent_id is not None
                        }
                        
                        # הוספה למיפוי ולרשימה
                        comment_map[comment_id] = comment_obj
                        all_comments.append(comment_obj)
                
                # *** שיפור: בניית עץ שרשורים מלא ***
                # מיון הערות לפי עיקריות ותגובות
                main_comments = []  # הערות ראשיות
                replies_by_parent = {}  # מיפוי של תגובות לפי מזהה הורה
                
                # חלוקה של ההערות להערות ראשיות ותגובות
                for comment in all_comments:
                    if not comment['parent_id']:  # הערה ראשית
                        main_comments.append(comment)
                    else:  # תגובה
                        parent_id = comment['parent_id']
                        if parent_id not in replies_by_parent:
                            replies_by_parent[parent_id] = []
                        replies_by_parent[parent_id].append(comment)
                
                # מיון הערות ראשיות לפי עמוד
                main_comments.sort(key=lambda x: x['page'])
                
                # ארגון השרשורים
                for main_comment in main_comments:
                    thread = {
                        'הערה': main_comment['text'],
                        'כותב': main_comment['author'],
                        'תאריך': main_comment['date'],
                        'עמוד': main_comment['page'],
                        'מזהה': main_comment['id'],
                    }
                    
                    # מציאת תגובות ישירות להערה הראשית
                    direct_replies = replies_by_parent.get(main_comment['id'], [])
                    
                    # מיון תגובות לפי תאריך (מהמוקדם למאוחר)
                    direct_replies.sort(key=lambda x: x.get('date', ''))
                    
                    # הוספת התגובות לשרשור (עד 5 תגובות)
                    for i, reply in enumerate(direct_replies[:5], 1):
                        thread[f'תגובה {i}'] = reply['text']
                        thread[f'כותב תגובה {i}'] = reply['author'] 
                        thread[f'תאריך תגובה {i}'] = reply['date']
                    
                    # הוספת השרשור לרשימה
                    comment_threads.append(thread)
                
                return comment_threads
                
        except Exception as e:
            print(f"שגיאה כללית בחילוץ הערות: {str(e)}")
            raise
    
    def calculate_page_numbers(self, doc, page_map):
        """חישוב מספרי עמודים לכל הערה (הערכה)"""
        # הערכה גסה של תוכן עמוד
        chars_per_page = 1500  # הערכה של מספר תווים בעמוד
        total_chars = 0
        current_page = 1
        
        # עבור על כל הפסקאות במסמך
        for para in doc.paragraphs:
            total_chars += len(para.text)
            estimated_page = max(1, (total_chars // chars_per_page) + 1)
            
            # חיפוש מזהי הערות בפסקה
            for run in para.runs:
                if hasattr(run, '_element') and run._element is not None:
                    comment_references = run._element.xpath(".//w:commentReference")
                    for ref in comment_references:
                        if ref is not None and ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id"):
                            comment_id = ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                            page_map[comment_id] = estimated_page
    
    def format_date(self, date_str):
        """עיצוב תאריך לפורמט קריא"""
        if not date_str:
            return ""
        
        try:
            # ניקוי וסטנדרטיזציה של הפורמט
            date_str = date_str.replace('Z', '+00:00')
            date_obj = datetime.fromisoformat(date_str)
            return date_obj.strftime('%d/%m/%Y %H:%M')
        except:
            return date_str
    
    def export_to_excel(self):
        """ייצוא הטבלה לקובץ אקסל"""
        if not self.comments_data:
            messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
            return
        
        # בחירת מיקום לשמירת הקובץ
        excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="שמור קובץ אקסל",
            initialfile=f"הערות_וורד_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        
        if not excel_path:
            return
        
        try:
            # המרה ל-DataFrame
            df = pd.DataFrame(self.comments_data)
            
            # הגדרת סדר העמודות בצורה נכונה לקריאה מימין לשמאל
            column_order = ['מס\'', 'עמוד', 'הערה', 'כותב', 'תאריך']
            
            # הוספת עמודות תגובה
            for i in range(1, 6):  # תמיכה בעד 5 תגובות
                for field in [f'תגובה {i}', f'כותב תגובה {i}', f'תאריך תגובה {i}']:
                    if field in df.columns:
                        column_order.append(field)
            
            # סינון לעמודות שקיימות בדאטה
            final_columns = [col for col in column_order if col in df.columns]
            
            # ייצוא לאקסל עם הגדרות RTL
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df[final_columns].to_excel(writer, index=False, sheet_name='הערות')
                
                # הגדרת כיוון RTL לגיליון
                worksheet = writer.sheets['הערות']
                worksheet.sheet_view.rightToLeft = True
                
                # התאמת רוחב עמודות באקסל
                for i, column in enumerate(final_columns):
                    max_length = df[column].astype(str).apply(len).max()
                    adjusted_width = max(10, min(50, max_length + 2))  # מינימום 10, מקסימום 50
                    col_letter = chr(65 + i)
                    worksheet.column_dimensions[col_letter].width = adjusted_width
            
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{excel_path}")
            self.status_var.set(f"הנתונים יוצאו בהצלחה: {os.path.basename(excel_path)}")
            
            # פתיחת התיקייה המכילה את הקובץ
            try:
                os.startfile(os.path.dirname(excel_path))
            except:
                pass
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בייצוא לאקסל:\n{str(e)}")

def main():
    root = tk.Tk()
    
    # הגדרות RTL ברמת האפליקציה
    try:
        root.tk_strictMotif(False)
    except:
        pass
    
    # יצירת האפליקציה
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
