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
        default_font = font.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=10)
        try:
            self.root.tk_strictMotif(False)
        except:
            pass
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", font=("Segoe UI", 10))
        style.configure("TLabel", anchor="e")
        style.configure("TButton", anchor="e")
        self.root.configure(bg="#f0f0f0")
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill="x", pady=10)
        
        header = ttk.Label(header_frame, text="ממיר הערות מקובץ וורד לאקסל", 
                          font=("Segoe UI", 16, "bold"))
        header.pack(side="right", padx=10)
        
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=10)
        
        select_btn = ttk.Button(file_frame, text="בחר קובץ וורד", command=self.select_file)
        select_btn.pack(side="right", padx=5)
        
        self.file_label = ttk.Label(file_frame, text="לא נבחר קובץ", anchor="e")
        self.file_label.pack(side="right", padx=5, fill="x", expand=True)
        
        process_frame = ttk.Frame(main_frame)
        process_frame.pack(fill="x", pady=5)
        
        process_btn = ttk.Button(process_frame, text="עבד את הקובץ", command=self.process_file)
        process_btn.pack(side="right", padx=5)
        
        status_frame = ttk.Frame(main_frame, relief="sunken", borderwidth=1)
        status_frame.pack(fill="x", pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("מוכן")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor="e")
        status_label.pack(side="right", pady=5, fill="x", expand=True)
        
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True, pady=10)
        
        columns = ("כותב 3", "תגובה 3", "כותב 2", "תגובה 2", "כותב 1", "תגובה 1", 
                   "תאריך", "עמוד", "כותב", "הערה", "#")
        
        self.result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        
        for col in columns:
            self.result_tree.heading(col, text=col, anchor="e")
            if col == "#":
                width = 40
            elif col == "עמוד":
                width = 50
            elif col == "תאריך":
                width = 120
            elif col == "כותב" or col.startswith("כותב"):
                width = 100
            elif col == "הערה" or col.startswith("תגובה"):
                width = 180
            else:
                width = 100
            self.result_tree.column(col, width=width, anchor="e", stretch=False)
        
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        y_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        y_scrollbar.pack(side="left", fill="y")
        self.result_tree.pack(side="left", fill="both", expand=True)
        x_scrollbar.pack(side="bottom", fill="x")
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        export_btn = ttk.Button(button_frame, text="ייצא לאקסל", command=self.export_to_excel)
        export_btn.pack(side="right", padx=5)
        
        exit_btn = ttk.Button(button_frame, text="יציאה", command=self.root.quit)
        exit_btn.pack(side="left", padx=5)
        
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
            self.result_tree.delete(*self.result_tree.get_children())
            self.comments_data = []
            comments_with_replies = self.extract_comments_with_replies()
            for idx, comment_thread in enumerate(comments_with_replies, 1):
                values = [
                    comment_thread.get('כותב 3', ''),
                    self.truncate_text(comment_thread.get('תגובה 3', ''), 50),
                    comment_thread.get('כותב 2', ''),
                    self.truncate_text(comment_thread.get('תגובה 2', ''), 50),
                    comment_thread.get('כותב 1', ''),
                    self.truncate_text(comment_thread.get('תגובה 1', ''), 50),
                    comment_thread.get('תאריך', ''),
                    comment_thread.get('עמוד', ''),
                    comment_thread.get('כותב', ''),
                    self.truncate_text(comment_thread.get('הערה', ''), 50),
                    idx
                ]
                self.result_tree.insert("", "end", values=values)
                comment_thread['אינדקס'] = idx
            self.comments_data = comments_with_replies
            if comments_with_replies:
                self.status_var.set(f"נמצאו {len(comments_with_replies)} הערות עם שרשורי תגובות. ניתן לייצא לאקסל.")
            else:
                self.status_var.set("לא נמצאו הערות בקובץ.")
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בעיבוד הקובץ:\n{str(e)}")
            self.status_var.set("אירעה שגיאה")
    
    def truncate_text(self, text, length=50):
        if not text:
            return ""
        return text[:length] + '...' if len(text) > length else text
    
    def extract_comments_with_replies(self):
        """
        חילוץ הערות ותגובות מוורד תוך זיהוי נכון של שרשור התגובות.
        במקרה שהתכונה 'parentId' אינה קיימת או ריקה, נבדוק את 'inReplyTo'.
        """
        if not self.docx_path or not os.path.exists(self.docx_path):
            return []
        try:
            result_threads = []
            with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
                comment_files = [f for f in docx_zip.namelist() if 'comments' in f or 'comment' in f]
                if not comment_files:
                    print("לא נמצאו קבצי הערות")
                    return []
                doc = Document(self.docx_path)
                total_chars = sum(len(p.text) for p in doc.paragraphs)
                pages_estimate = max(1, total_chars // 1800)
                all_comments = []
                for comment_file in comment_files:
                    try:
                        xml_content = docx_zip.read(comment_file)
                        root = ET.fromstring(xml_content)
                        ns_match = re.search(r'{(.*?)}', root.tag)
                        if not ns_match:
                            continue
                        ns = {'w': ns_match.group(1)}
                        for comment in root.findall('.//w:comment', ns):
                            comment_id = comment.get(f"{{{ns['w']}}}id")
                            if not comment_id:
                                continue
                            # ננסה לקרוא קודם את parentId ואם אין, נבדוק inReplyTo
                            parent_id = comment.get(f"{{{ns['w']}}}parentId", "").strip()
                            if not parent_id:
                                parent_id = comment.get(f"{{{ns['w']}}}inReplyTo", "").strip()
                            author = comment.get(f"{{{ns['w']}}}author", "")
                            date = comment.get(f"{{{ns['w']}}}date", "")
                            text_parts = []
                            for text_elem in comment.findall('.//w:t', ns):
                                if text_elem.text:
                                    text_parts.append(text_elem.text)
                            comment_text = "".join(text_parts)
                            page = 1
                            if len(all_comments) > 0:
                                position_ratio = len(all_comments) / 20
                                page = int(1 + (position_ratio * pages_estimate))
                                page = min(pages_estimate, max(1, page))
                            all_comments.append({
                                'id': comment_id,
                                'parent_id': parent_id,
                                'author': author,
                                'date': self.format_date(date),
                                'text': comment_text,
                                'page': page
                            })
                    except Exception as e:
                        print(f"שגיאה בעיבוד קובץ {comment_file}: {str(e)}")
                print(f"נמצאו {len(all_comments)} הערות במסמך")
                main_comments = []
                reply_map = {}
                for comment in all_comments:
                    # אם אין ערך ב-parent_id – זו הערה ראשית, אחרת תגובה
                    if not comment['parent_id']:
                        main_comments.append(comment)
                    else:
                        parent_id = comment['parent_id']
                        if parent_id not in reply_map:
                            reply_map[parent_id] = []
                        reply_map[parent_id].append(comment)
                print(f"מתוכן {len(main_comments)} הערות ראשיות")
                for main_comment in main_comments:
                    comment_id = main_comment['id']
                    thread = {
                        'הערה': main_comment['text'],
                        'כותב': main_comment['author'],
                        'תאריך': main_comment['date'],
                        'עמוד': main_comment['page']
                    }
                    direct_replies = reply_map.get(comment_id, [])
                    direct_replies.sort(key=lambda x: x['date'])
                    for i, reply in enumerate(direct_replies, 1):
                        if i <= 10:
                            thread[f'תגובה {i}'] = reply['text']
                            thread[f'כותב {i}'] = reply['author']
                    result_threads.append(thread)
                result_threads.sort(key=lambda x: x.get('עמוד', 0))
                return result_threads
        except Exception as e:
            print(f"שגיאה כללית בחילוץ הערות: {str(e)}")
            raise e
    
    def format_date(self, date_str):
        if not date_str:
            return ""
        try:
            date_str = date_str.replace('Z', '+00:00')
            date_obj = datetime.fromisoformat(date_str)
            return date_obj.strftime('%d/%m/%Y %H:%M')
        except:
            return date_str
    
    def export_to_excel(self):
        if not self.comments_data:
            messagebox.showwarning("אזהרה", "אין נתונים לייצוא")
            return
        excel_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="שמור קובץ אקסל",
            initialfile=f"הערות_וורד_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
        if not excel_path:
            return
        try:
            df = pd.DataFrame(self.comments_data)
            column_order = ['אינדקס', 'הערה', 'כותב', 'עמוד', 'תאריך']
            for i in range(1, 11):
                for field in [f'תגובה {i}', f'כותב {i}']:
                    if field in df.columns:
                        column_order.append(field)
            final_columns = [col for col in column_order if col in df.columns]
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df[final_columns].to_excel(writer, index=False, sheet_name='הערות')
                try:
                    worksheet = writer.sheets['הערות']
                    worksheet.sheet_view.rightToLeft = True
                except:
                    pass
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{excel_path}")
            self.status_var.set("הנתונים יוצאו בהצלחה")
            try:
                os.startfile(os.path.dirname(excel_path))
            except:
                pass
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה בייצוא לאקסל:\n{str(e)}")

def main():
    root = tk.Tk()
    try:
        root.tk_strictMotif(False)
    except:
        pass
    app = WordCommentsExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
