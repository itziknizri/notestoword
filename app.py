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
        self.root.geometry("900x650")
        
        # הגדרת ערכי התצוגה
        self.docx_path = None
        self.comments_data = []
        
        # הגדרת כיוון RTL
        self.configure_rtl_support()
        
        # יצירת ממשק
        self.create_widgets()
        
        # שיפור המראה
        self.style_widgets()
    
    def configure_rtl_support(self):
        """הגדרת תמיכה בכיוון מימין לשמאל"""
        # יצירת פונטים
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        
        # הגדרת RTL באופן מערכתי
        try:
            # הגדרות בסיסיות
            self.root.tk_strictMotif(False)
        except:
            pass
        
        # הגדרת סגנון טקסט לימין
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10))
        style.configure("TLabel", anchor="e")  # טקסט לימין
        style.configure("TButton", anchor="e")  # טקסט לימין
    
    def style_widgets(self):
        # עיצוב הממשק
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
        
        # הגדרת עמודות בסדר הנכון (מימין לשמאל!)
        columns = ("תאריך תגובה 3", "כותב תגובה 3", "תגובה 3", 
                  "תאריך תגובה 2", "כותב תגובה 2", "תגובה 2", 
                  "תאריך תגובה 1", "כותב תגובה 1", "תגובה 1", 
                  "תאריך", "עמוד", "כותב", "הערה", "#")
        
        # יצירת טבלה עם סדר עמודות הפוך
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # הגדרת כותרות
        for col in columns:
            self.result_tree.heading(col, text=col, anchor="e")  # כותרות מיושרות לימין
            width = 80 if "תאריך" in col else 120 if "תגובה" in col else 80 if "כותב" in col else 40 if col == "#" else 50 if col == "עמוד" else 150
            self.result_tree.column(col, width=width, anchor="e")  # תוכן מיושר לימין
        
        # הגדרת סרגלי גלילה
        y_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # סידור רכיבי הטבלה
        y_scrollbar.pack(side="left", fill="y")  # שים לב: סרגל גלילה אנכי בצד שמאל
        x_scrollbar.pack(side="bottom", fill="x")
        self.result_tree.pack(side="right", fill="both", expand=True)  # טבלה בצד ימין
        
        # מסגרת כפתורים
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        # כפתור ייצוא
        export_btn = ttk.Button(button_frame, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(side="right", padx=5)
        
        # כפתור יציאה
        exit_btn = ttk.Button(button_frame, text="יציאה", command=self.root.quit)
        exit_btn.pack(side="left", padx=5)
        
        # קרדיט (מיושר למרכז)
        credit_frame = ttk.Frame(main_frame)
        credit_frame.pack(fill="x", pady=5)
        
        credit = ttk.Label(credit_frame, text="פותח ע\"י יצחק נזרי", 
                          font=("Segoe UI", 9), foreground="gray")
        credit.pack(side="bottom", pady=5)
    
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
            # ניקוי טבלה קודמת
            self.result_tree.delete(*self.result_tree.get_children())
            self.comments_data = []
            
            # חילוץ ההערות
            comments = self.extract_comments_from_docx(self.docx_path)
            
            # הצגת התוצאות בטבלה
            for idx, comment in enumerate(comments, 1):
                # כאן השינוי המשמעותי - סדר הערכים מותאם לסדר העמודות מימין לשמאל
                values = [
                    comment.get('תאריך תגובה 3', ''),
                    comment.get('כותב תגובה 3', ''),
                    self.truncate_text(comment.get('תגובה 3', ''), 50),
                    comment.get('תאריך תגובה 2', ''),
                    comment.get('כותב תגובה 2', ''),
                    self.truncate_text(comment.get('תגובה 2', ''), 50),
                    comment.get('תאריך תגובה 1', ''),
                    comment.get('כותב תגובה 1', ''),
                    self.truncate_text(comment.get('תגובה 1', ''), 50),
                    comment.get('תאריך', ''),
                    comment.get('עמוד', ''),
                    comment.get('כותב', ''),
                    self.truncate_text(comment.get('הערה', ''), 50),
                    idx  # מספר הערה
                ]
                self.result_tree.insert("", "end", values=values)
                
                # שמירת מספר האינדקס בנתונים
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
        פונקציה משופרת לחילוץ הערות ושרשורי תגובות מקובץ וורד
        """
        try:
            # קריאת המסמך
            doc = Document(docx_path)
            
            # אומדן מספר העמודים 
            total_chars = sum(len(p.text) for p in doc.paragraphs)
            chars_per_page = 1800  # אומדן תווים לעמוד
            estimated_pages = max(1, total_chars // chars_per_page)
            
            # פתיחת הקובץ כארכיון ZIP
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                # חיפוש קובץ ההערות
                comment_files = [f for f in zip_ref.namelist() if 'comments.xml' in f or 'comment' in f]
                
                if not comment_files:
                    return []
                
                # נאסוף את כל ההערות והתגובות
                comments_and_replies = []
                
                for comment_file in comment_files:
                    xml_content = zip_ref.read(comment_file)
                    try:
                        root = ET.fromstring(xml_content)
                    except ET.ParseError:
                        continue  # נדלג על קבצים לא תקינים
                    
                    # איתור namespace
                    ns_match = re.search(r'{(.*?)}', root.tag)
                    if not ns_match:
                        continue
                        
                    ns = {'w': ns_match.group(1)}
                    
                    # איסוף כל ההערות מה-XML
                    comment_elements = root.findall('.//w:comment', ns)
                    
                    for comment in comment_elements:
                        comment_id = comment.get(f"{{{ns['w']}}}id")
                        author = comment.get(f"{{{ns['w']}}}author", "")
                        date = comment.get(f"{{{ns['w']}}}date", "")
                        parent_id = comment.get(f"{{{ns['w']}}}parentId")
                        
                        # חילוץ טקסט ההערה
                        paragraphs = comment.findall('.//w:p', ns)
                        comment_text = ""
                        
                        for p in paragraphs:
                            runs = p.findall('.//w:r', ns)
                            for r in runs:
                                text_elements = r.findall('.//w:t', ns)
                                for t in text_elements:
                                    if t.text:
                                        comment_text += t.text
                        
                        # חישוב מיקום משוער בטקסט => מספר עמוד
                        # הערה: בהיעדר API מדויק, נחלק את המסמך לחלקים שווים
                        relative_position = 0
                        if len(comments_and_replies) > 0:
                            relative_position = len(comments_and_replies) / 20  # הערכה גסה
                        
                        page = max(1, min(estimated_pages, 
                                         int(1 + (relative_position * estimated_pages))))
                        
                        comments_and_replies.append({
                            'id': comment_id,
                            'parent_id': parent_id,
                            'text': comment_text,
                            'author': author,
                            'date': self.format_date(date),
                            'page': page
                        })
                
                # ארגון ההערות ותגובותיהן
                comment_threads = {}
                
                # סינון להערות ראשיות (ללא parent_id)
                main_comments = [c for c in comments_and_replies if c['parent_id'] is None]
                
                # מיפוי תגובות להערות
                for comment in main_comments:
                    # מציאת כל התגובות להערה הזו
                    replies = [r for r in comments_and_replies 
                              if r['parent_id'] == comment['id']]
                    
                    # מיון התגובות לפי תאריך
                    replies.sort(key=lambda x: x.get('date', ''))
                    
                    # יצירת שורה עבור ההערה עם כל התגובות
                    comment_row = {
                        'אינדקס': len(comment_threads) + 1,
                        'הערה': comment['text'],
                        'כותב': comment['author'],
                        'תאריך': comment['date'],
                        'עמוד': comment['page']
                    }
                    
                    # הוספת התגובות כעמודות נפרדות
                    for i, reply in enumerate(replies, 1):
                        if i <= 10:  # תמיכה בעד 10 תגובות
                            comment_row[f'תגובה {i}'] = reply['text']
                            comment_row[f'כותב תגובה {i}'] = reply['author']
                            comment_row[f'תאריך תגובה {i}'] = reply['date']
                    
                    comment_threads[comment['id']] = comment_row
                
                # המרה למערך תוצאות
                result = list(comment_threads.values())
                
                # מיון לפי מספר עמוד
                result.sort(key=lambda x: x.get('עמוד', 0))
                
                return result
                
        except Exception as e:
            print(f"שגיאה בחילוץ ההערות: {str(e)}")
            raise e
    
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
        if not self.comments_data:
            messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
            return
        
        # בחירת מיקום וייצוא
        excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="שמור קובץ אקסל",
            initialfile=f"הערות_וורד_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        
        if not excel_path:
            return
        
        try:
            # יצירת DataFrame במבנה הנכון
            df = pd.DataFrame(self.comments_data)
            
            # הגדרת סדר העמודות בצורה נכונה לקריאה מימין לשמאל
            column_order = ['אינדקס', 'הערה', 'כותב', 'עמוד', 'תאריך']
            
            # הוספת עמודות תגובה
            for i in range(1, 11):
                for field in [f'תגובה {i}', f'כותב תגובה {i}', f'תאריך תגובה {i}']:
                    if field in df.columns:
                        column_order.append(field)
            
            # סינון לעמודות שקיימות בדאטה
            final_columns = [col for col in column_order if col in df.columns]
            
            # ייצוא לאקסל
            df = df[final_columns]
            
            # הגדרת כיוון RTL לאקסל
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='הערות')
                # הגדרת הגיליון לRTL
                writer.sheets['הערות'].sheet_view.rightToLeft = True
            
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{excel_path}")
            self.status_var.set("הנתונים יוצאו בהצלחה")
            
            # ניסיון לפתוח את התיקייה 
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
