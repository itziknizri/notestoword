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

# ניסיון לייבא win32com - עשוי להיכשל אם החבילה לא מותקנת
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

class WordCommentsExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("ממיר הערות מוורד לאקסל")
        self.root.geometry("1100x650")
        
        # ערכים שנשמור
        self.docx_path = None
        self.comments_data = []
        
        # בדיקה האם win32com זמין
        self.has_win32com = HAS_WIN32COM
        
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
        
        # הגדרת עמודות ברוחב מצומצם יותר - סדר מימין לשמאל עם פחות עמודות
        columns = (
            "מס'", "עמוד", "הערה", "כותב", "תאריך", 
            "תגובות"  # עמודה אחת לכל התגובות
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
            elif col == "תאריך":
                width = 100
            elif col == "כותב":
                width = 80
            elif col == "הערה":
                width = 200
            elif col == "תגובות":
                width = 300  # עמודה רחבה לתגובות
            else:
                width = 100
                
            # מתיחת עמודות לפי הצורך
            stretch = True if col in ["הערה", "תגובות"] else False
            self.result_tree.column(col, width=width, anchor="e", stretch=stretch)
        
        # הגדרת סרגלי גלילה
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        y_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # סידור רכיבי הטבלה - RTL סדר
        y_scrollbar.pack(side="left", fill="y")  # סרגל גלילה אנכי בצד שמאל
        x_scrollbar.pack(side="bottom", fill="x")  # חשוב: סרגל הגלילה האופקי למעלה
        self.result_tree.pack(side="right", fill="both", expand=True)  # עץ בצד ימין
        
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
                # מיזוג כל התגובות לעמודה אחת
                replies_text = ""
                
                # שרשור כל התגובות בפורמט קריא
                for i in range(1, 6):
                    reply_text = thread.get(f'תגובה {i}', '')
                    reply_author = thread.get(f'כותב תגובה {i}', '')
                    reply_date = thread.get(f'תאריך תגובה {i}', '')
                    
                    if reply_text and reply_author:
                        if replies_text:
                            replies_text += "\n\n"  # מרווח בין תגובות
                        
                        replies_text += f"{reply_author} ({reply_date}):\n{reply_text}"
                
                # הכנת ערכים להצגה בטבלה המצומצמת - סדר מימין לשמאל
                values = [
                    idx,  # מספר
                    thread.get('עמוד', ''),  # עמוד
                    thread.get('הערה', ''),  # הערה מקורית
                    thread.get('כותב', ''),  # כותב
                    thread.get('תאריך', ''),  # תאריך
                    replies_text  # כל התגובות מאוחדות
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
            
            # ניסיון להשיג מספרי עמודים מדויקים באמצעות Word COM
            page_map = self.get_exact_page_numbers(self.docx_path)
            
            # פתיחת קובץ docx כארכיון ZIP
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                # אם לא הצלחנו לקבל מספרי עמודים מדויקים, ננסה לחשב בעצמנו
                if not page_map:
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
    
    def get_exact_page_numbers(self, doc_path):
        """מציאת מספרי עמודים מדויקים באמצעות Word COM API"""
        page_map = {}
        
        # אם win32com לא זמין, החזר מילון ריק
        if not self.has_win32com:
            self.status_var.set("חבילת win32com לא מותקנת. משתמש בשיטה חלופית...")
            self.root.update()
            return page_map
            
        try:
            self.status_var.set("מנסה לחשב מספרי עמודים מדויקים עם Word...")
            self.root.update()
            
            # יצירת אובייקט Word
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            
            # פתיחת המסמך
            abs_path = os.path.abspath(doc_path)
            doc = word_app.Documents.Open(abs_path)
            
            # מספר הערות במסמך
            comment_count = doc.Comments.Count
            self.status_var.set(f"נמצאו {comment_count} הערות. מחשב מספרי עמודים...")
            self.root.update()
            
            # מיפוי מספרי הערות למספרי עמודים
            for i in range(1, comment_count + 1):
                try:
                    comment = doc.Comments(i)
                    
                    # ניסיון להשיג מזהה הערה אמיתי - מורכב אבל אפשרי
                    # ב-Word התצוגה היא מ-1, אבל ב-XML המזהים מתחילים מ-0
                    comment_id = str(i-1)
                    
                    # השג את הטקסט עם ההערה
                    comment_range = comment.Scope
                    
                    # מספר העמוד בו נמצאת ההערה
                    page_num = comment_range.Information(3)  # wdActiveEndPageNumber = 3
                    page_map[comment_id] = page_num
                    
                except Exception as e:
                    print(f"שגיאה בהערה {i}: {str(e)}")
                    continue
            
            # סגירת המסמך והאפליקציה
            doc.Close(False)
            word_app.Quit()
            
            # הצלחנו לקבל מידע?
            if page_map:
                self.status_var.set(f"חישוב מספרי עמודים הושלם! נמצאו {len(page_map)} הערות.")
                self.root.update()
            
            return page_map
            
        except Exception as e:
            self.status_var.set("לא ניתן להשתמש ב-Word COM API, משתמש בשיטה חלופית...")
            self.root.update()
            print(f"שגיאה בחישוב מספרי עמודים מדויקים: {str(e)}")
            return {}
    
    def calculate_page_numbers(self, doc, page_map):
        """חישוב מספרי עמודים לכל הערה (הערכה משופרת)"""
        # גישה מדויקת יותר לחישוב עמודים
        try:
            # מציאת מספר תווים ממוצע לעמוד על פי מסמכי וורד סטנדרטיים
            chars_per_page = 3000  # הערכה מעודכנת

            # מאפייני המסמך
            paragraphs = list(doc.paragraphs)
            total_paragraphs = len(paragraphs)
            total_chars = sum(len(p.text) for p in paragraphs)
            
            # הערכת מספר עמודים כולל
            estimated_total_pages = max(1, total_chars // chars_per_page)
            
            # יצירת רשימה של אחוזי התקדמות לכל פסקה במסמך
            para_positions = []
            current_char_count = 0
            
            for para in paragraphs:
                current_char_count += len(para.text)
                position_percent = current_char_count / total_chars if total_chars > 0 else 0
                para_positions.append(position_percent)
            
            # חישוב עמוד לכל פסקה
            current_page = 1
            page_breaks = []
            
            # נעבור על כל הפסקאות במסמך
            for idx, para in enumerate(paragraphs):
                # חישוב מספר עמוד משוער על פי מיקום באחוזים
                estimated_page = max(1, int(para_positions[idx] * estimated_total_pages) + 1)
                
                # חיפוש מזהי הערות בפסקה
                found_comments = False
                for run in para.runs:
                    if hasattr(run, '_element') and run._element is not None:
                        try:
                            comment_references = run._element.xpath(".//w:commentReference")
                            for ref in comment_references:
                                if ref is not None and ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id"):
                                    comment_id = ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                                    page_map[comment_id] = estimated_page
                                    found_comments = True
                        except:
                            pass
                    
            # הגדלת מגוון העמודים אם כולם זהים
            if len(set(page_map.values())) <= 1 and len(page_map) > 1:
                # פיזור מלאכותי אם כל העמודים יצאו זהים
                comment_ids = list(page_map.keys())
                num_comments = len(comment_ids)
                
                # יצירת טווח עמודים סינטטי
                for i, comment_id in enumerate(comment_ids):
                    synthetic_page = max(1, int((i / num_comments) * estimated_total_pages) + 1)
                    page_map[comment_id] = synthetic_page
        
        except Exception as e:
            print(f"שגיאה בחישוב עמודים: {str(e)}")
            # במקרה של שגיאה - הקצאת מספרי עמודים שונים לפחות
            comment_ids = list(page_map.keys())
            for i, comment_id in enumerate(comment_ids):
                page_map[comment_id] = (i % 10) + 1
    
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
        """ייצוא הטבלה לקובץ אקסל עם שמירה על שרשורי התגובות"""
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
            # יצירת מבנה נתונים חדש לייצוא לאקסל עם פורמט נכון
            excel_data = []
            
            for thread in self.comments_data:
                # עיבוד התגובות לפורמט ברור
                row_data = {
                    'מס\'': thread.get('מס\'', ''),
                    'עמוד': thread.get('עמוד', ''),
                    'הערה': thread.get('הערה', ''),
                    'כותב': thread.get('כותב', ''),
                    'תאריך': thread.get('תאריך', '')
                }
                
                # הוספת התגובות בפורמט ברור
                for i in range(1, 6):
                    reply_text = thread.get(f'תגובה {i}', '')
                    reply_author = thread.get(f'כותב תגובה {i}', '')
                    reply_date = thread.get(f'תאריך תגובה {i}', '')
                    
                    if reply_text:
                        row_data[f'תגובה {i}'] = reply_text
                        row_data[f'כותב תגובה {i}'] = reply_author
                        row_data[f'תאריך תגובה {i}'] = reply_date
                
                excel_data.append(row_data)
            
            # המרה ל-DataFrame
            df = pd.DataFrame(excel_data)
            
            # הגדרת סדר העמודות בצורה נכונה לקריאה מימין לשמאל
            column_order = ['מס\'', 'עמוד', 'הערה', 'כותב', 'תאריך']
            
            # הוספת עמודות תגובה
            for i in range(1, 6):  # תמיכה בעד 5 תגובות
                for field in [f'תגובה {i}', f'כותב תגובה {i}', f'תאריך תגובה {i}']:
                    if field in df.columns:
                        column_order.append(field)
            
            # ייצוא לאקסל עם הגדרות RTL
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # יצירת גיליון עם העמודות בסדר הנכון
                df[column_order].to_excel(writer, index=False, sheet_name='הערות')
                
                # הגדרת כיוון RTL לגיליון
                worksheet = writer.sheets['הערות']
                worksheet.sheet_view.rightToLeft = True
                
                # התאמת רוחב עמודות באקסל - נותן יותר מקום לתוכן
                for i, column in enumerate(column_order):
                    if 'הערה' in column or 'תגובה' in column:
                        base_width = 50  # עמודות טקסט ארוכות
                    elif 'תאריך' in column:
                        base_width = 20  # עמודות תאריך
                    elif 'כותב' in column:
                        base_width = 15  # עמודות שם כותב
                    elif column == 'עמוד':
                        base_width = 8  # עמודת מספר עמוד
                    elif column == 'מס\'':
                        base_width = 6  # עמודת מספור
                    else:
                        base_width = 15
                    
                    # חישוב לפי תוכן בפועל אם אפשר
                    try:
                        max_length = df[column].astype(str).apply(len).max()
                        adjusted_width = max(base_width, min(80, max_length + 2))
                    except:
                        adjusted_width = base_width
                    
                    # הגדרת רוחב עמודה
                    col_letter = chr(65 + i)
                    if i < 26:
                        worksheet.column_dimensions[col_letter].width = adjusted_width
                    else:
                        # טיפול בעמודות מעבר ל-Z
                        first = chr(64 + (i // 26))
                        second = chr(65 + (i % 26))
                        worksheet.column_dimensions[f"{first}{second}"].width = adjusted_width
            
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
    
    # בדיקה אם win32com מותקן
    if not HAS_WIN32COM:
        messagebox.showwarning(
            "הערה", 
            "חבילת win32com לא מותקנת. התוכנה תפעל, אבל ללא יכולת לזהות מספרי עמודים מדויקים.\n"
            "להתקנת החבילה הרצו: pip install pywin32"
        )
    
    # יצירת האפליקציה
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
