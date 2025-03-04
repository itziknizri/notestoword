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
        
        # הגדרת עמודות בסדר הנכון (מימין לשמאל!)
        columns = ("תאריך", "עמוד", "כותב", "הערה", "#")
        
        # יצירת טבלה עם סדר עמודות מימין לשמאל
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        
        # הגדרת כותרות
        for col in columns:
            self.result_tree.heading(col, text=col, anchor="e")  # כותרות מיושרות לימין
            
            # רוחב קבוע לעמודות
            if col == "#":
                width = 40
            elif col == "עמוד":
                width = 50
            elif col == "תאריך":
                width = 120
            elif col == "כותב":
                width = 120
            elif col == "הערה":
                width = 300
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
            
            # חילוץ כל ההערות (כולל תגובות) כרשימה שטוחה
            all_comments = self.extract_all_comments()
            
            # הצגת התוצאות בטבלה
            for idx, comment in enumerate(all_comments, 1):
                # מכין ערכים להצגה בטבלה
                values = [
                    comment.get('תאריך', ''),
                    comment.get('עמוד', ''),
                    comment.get('כותב', ''),
                    self.format_comment_text(comment),
                    idx
                ]
                
                # קובע אם זו הערה ראשית או תגובה לצורך העיצוב
                is_reply = comment.get('parent_id') is not None
                tag = "reply" if is_reply else "main"
                
                # מוסיף לטבלה
                self.result_tree.insert("", "end", values=values, tags=(tag,))
                
                # שומר את האינדקס למידע
                comment['אינדקס'] = idx
            
            # הגדרת צבעים להערות ותגובות
            self.result_tree.tag_configure("main", background="white")
            self.result_tree.tag_configure("reply", background="#e6f2ff")  # כחול בהיר לתגובות
            
            self.comments_data = all_comments
            
            if all_comments:
                self.status_var.set(f"נמצאו {len(all_comments)} הערות ותגובות. ניתן לייצא לאקסל.")
            else:
                self.status_var.set("לא נמצאו הערות בקובץ.")
        
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בעיבוד הקובץ:\n{str(e)}")
            self.status_var.set("אירעה שגיאה")
    
    def format_comment_text(self, comment):
        """מעצב טקסט הערה או תגובה עם סימון מתאים"""
        text = comment.get('text', '')
        if text:
            # אם זה תגובה, הוסף סימון לפניה
            if comment.get('parent_id') is not None:
                text = f"↪️ {text}"  # חץ מעוגל לציון תגובה
        return self.truncate_text(text, 200)
    
    def truncate_text(self, text, length=50):
        """קיצור טקסט ארוך לתצוגה בטבלה"""
        if not text:
            return ""
        return text[:length] + '...' if len(text) > length else text
    
    def extract_all_comments(self):
        """
        מחלץ את כל ההערות והתגובות מקובץ וורד ברשימה שטוחה
        עם סימון הקשרים ביניהן
        """
        if not self.docx_path or not os.path.exists(self.docx_path):
            return []
            
        try:
            all_comments = []  # הרשימה הסופית של כל ההערות והתגובות
            
            # קריאת המסמך להערכת מספרי עמודים
            doc = Document(self.docx_path)
            total_chars = sum(len(p.text) for p in doc.paragraphs)
            pages_estimate = max(1, total_chars // 1800)  # הערכה של מספר עמודים
            
            # פתיחת קובץ docx כארכיון ZIP
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                # חיפוש קבצי XML שמכילים הערות
                comment_files = [f for f in docx_zip.namelist() if 'comments' in f or 'comment' in f]
                
                if not comment_files:
                    return []
                
                # מיפוי של מזהי הערות להערות עצמן (לשימוש מאוחר יותר)
                comment_map = {}
                
                # קריאת כל קבצי ההערות
                for comment_file in comment_files:
                    try:
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
                                
                            parent_id = comment.get(f"{{{ns['w']}}}parentId")
                            author = comment.get(f"{{{ns['w']}}}author", "")
                            date = comment.get(f"{{{ns['w']}}}date", "")
                            
                            # חילוץ טקסט ההערה
                            text_parts = []
                            for text_elem in comment.findall('.//w:t', ns):
                                if text_elem.text:
                                    text_parts.append(text_elem.text)
                            
                            comment_text = "".join(text_parts)
                            
                            # חישוב מספר עמוד משוער
                            # ככל שההערה מאוחרת יותר, כך היא כנראה בעמוד מאוחר יותר
                            page = 1
                            if len(all_comments) > 0:
                                position_ratio = len(all_comments) / 20  # הערכה גסה
                                page = int(1 + (position_ratio * pages_estimate))
                                page = min(pages_estimate, max(1, page))
                            
                            # יצירת אובייקט ההערה
                            comment_obj = {
                                'id': comment_id,
                                'parent_id': parent_id,
                                'author': author,
                                'date': self.format_date(date),
                                'text': comment_text,
                                'page': page,
                                'is_reply': parent_id is not None  # סימון האם זו תגובה
                            }
                            
                            # הוספה למיפוי לשימוש מאוחר יותר
                            comment_map[comment_id] = comment_obj
                            all_comments.append(comment_obj)
                    
                    except Exception as e:
                        print(f"שגיאה בעיבוד קובץ {comment_file}: {str(e)}")
                
                # מיון הערות לפי קשרי הורה-ילד
                # הערות ראשיות קודם, ואז התגובות אליהן
                sorted_comments = []
                main_comments = [c for c in all_comments if c['parent_id'] is None]
                
                # הוספת הערות ראשיות
                for main in main_comments:
                    sorted_comments.append(main)
                    
                    # חיפוש תגובות ישירות להערה זו
                    direct_replies = [c for c in all_comments if c['parent_id'] == main['id']]
                    direct_replies.sort(key=lambda x: x.get('date', ''))
                    
                    # הוספת תגובות ישירות
                    for reply in direct_replies:
                        sorted_comments.append(reply)
                
                return sorted_comments
                
        except Exception as e:
            print(f"שגיאה כללית בחילוץ הערות: {str(e)}")
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
            # הכנת הנתונים לייצוא, עם התאמות לתגובות
            export_data = []
            for comment in self.comments_data:
                export_row = {
                    'אינדקס': comment.get('אינדקס', ''),
                    'סוג': 'תגובה' if comment.get('is_reply') else 'הערה',
                    'הערה': comment.get('text', ''),
                    'כותב': comment.get('author', ''),
                    'עמוד': comment.get('page', ''),
                    'תאריך': comment.get('date', ''),
                    'מזהה הורה': comment.get('parent_id', '')
                }
                export_data.append(export_row)
            
            # המרה ל-DataFrame
            df = pd.DataFrame(export_data)
            
            # ייצוא לאקסל עם הגדרות RTL
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='הערות')
                
                # הגדרת כיוון RTL לגיליון
                try:
                    worksheet = writer.sheets['הערות']
                    worksheet.sheet_view.rightToLeft = True
                    
                    # הרחבת עמודות באקסל לקריאות טובה יותר
                    for i, column in enumerate(df.columns):
                        column_width = max(len(str(column)), df[column].astype(str).map(len).max())
                        # הגבלת רוחב מקסימלי
                        column_width = min(column_width, 50)
                        worksheet.column_dimensions[chr(65 + i)].width = column_width + 2
                        
                except:
                    # במקרה שיש תקלה בהגדרת RTL
                    pass
            
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{excel_path}")
            self.status_var.set("הנתונים יוצאו בהצלחה")
            
            # ניסיון לפתוח את התיקייה המכילה את הקובץ
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
