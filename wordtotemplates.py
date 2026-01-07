#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
محوّل نماذج الوورد إلى Templates Data
تطوير: عبدالكريم العبود | abo.saleh.g@gmail.com
"""

import sys
import json
import os
from pathlib import Path

try:
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QLabel, QPushButton, QLineEdit, QTextEdit, QComboBox, QListWidget,
        QListWidgetItem, QFileDialog, QMessageBox, QSplitter, QGroupBox,
        QFormLayout, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
        QStyle, QStyleFactory, QInputDialog, QMenu, QAction, QStatusBar,
        QProgressBar, QFrame, QSpinBox, QCheckBox
    )
    from PyQt5.QtCore import Qt, QSize, QTimer
    from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QClipboard
except ImportError:
    print("يجب تثبيت PyQt5 أولاً:")
    print("pip install PyQt5")
    sys.exit(1)

try:
    from docx import Document
    from docx.table import Table
except ImportError:
    print("يجب تثبيت python-docx أولاً:")
    print("pip install python-docx")
    sys.exit(1)


class Template:
    """كائن النموذج"""
    def __init__(self, num='', keyword='', content='', category=''):
        self.num = num
        self.keyword = keyword
        self.content = content
        self.category = category
    
    def to_dict(self):
        return {
            'num': self.num,
            'keyword': self.keyword,
            'content': self.content
        }
    
    def __str__(self):
        return f"[{self.num}] {self.keyword}: {self.content[:50]}..."


class TemplateConverter(QMainWindow):
    """النافذة الرئيسية للتطبيق"""
    
    # التصنيفات الافتراضية
    DEFAULT_CATEGORIES = [
        'الدعوى', 'الإجابة', 'المرافعة', 'الأسباب', 'الحكم',
        'الشهادة', 'الصلح', 'اليمين', 'النكول', 'الكفالة',
        'الالتماس', 'الشطب', 'الغياب', 'الاختصاص', 'التمويل',
        'العقارات', 'المشاكل_التقنية'
    ]
    
    def __init__(self):
        super().__init__()
        self.templates = {}  # {category: [Template, ...]}
        self.current_file = None
        self.is_dark_mode = False
        self.init_categories()
        self.init_ui()
        self.apply_light_theme()
    
    def init_categories(self):
        """تهيئة التصنيفات"""
        for cat in self.DEFAULT_CATEGORIES:
            self.templates[cat] = []
    
    def init_ui(self):
        """إنشاء واجهة المستخدم"""
        self.setWindowTitle('محوّل نماذج الوورد إلى Templates Data')
        self.setMinimumSize(1200, 800)
        self.setLayoutDirection(Qt.RightToLeft)
        
        # الودجت المركزي
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # شريط الأدوات العلوي
        toolbar = self.create_toolbar()
        main_layout.addLayout(toolbar)
        
        # التبويبات
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_import_tab(), '📥 استيراد من وورد')
        self.tabs.addTab(self.create_manual_tab(), '✏️ إضافة يدوية')
        self.tabs.addTab(self.create_preview_tab(), '👁️ معاينة')
        self.tabs.addTab(self.create_export_tab(), '📤 تصدير')
        main_layout.addWidget(self.tabs)
        
        # شريط الحالة
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status()
        
        # شريط الحقوق
        credits = QLabel('تطوير: عبدالكريم العبود | abo.saleh.g@gmail.com')
        credits.setAlignment(Qt.AlignCenter)
        credits.setStyleSheet('color: #888; padding: 5px;')
        main_layout.addWidget(credits)
    
    def create_toolbar(self):
        """إنشاء شريط الأدوات"""
        layout = QHBoxLayout()
        
        # زر فتح ملف
        btn_open = QPushButton('📂 فتح ملف وورد')
        btn_open.clicked.connect(self.open_word_file)
        layout.addWidget(btn_open)
        
        # زر حفظ المشروع
        btn_save = QPushButton('💾 حفظ المشروع')
        btn_save.clicked.connect(self.save_project)
        layout.addWidget(btn_save)
        
        # زر تحميل المشروع
        btn_load = QPushButton('📁 تحميل مشروع')
        btn_load.clicked.connect(self.load_project)
        layout.addWidget(btn_load)
        
        layout.addStretch()
        
        # زر الوضع الليلي
        self.btn_theme = QPushButton('🌙 الوضع الليلي')
        self.btn_theme.clicked.connect(self.toggle_theme)
        layout.addWidget(self.btn_theme)
        
        return layout
    
    def create_import_tab(self):
        """تبويب الاستيراد من وورد"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # معلومات الملف
        file_group = QGroupBox('ملف الوورد')
        file_layout = QHBoxLayout(file_group)
        
        self.lbl_file = QLabel('لم يتم اختيار ملف')
        file_layout.addWidget(self.lbl_file)
        
        btn_browse = QPushButton('استعراض...')
        btn_browse.clicked.connect(self.open_word_file)
        file_layout.addWidget(btn_browse)
        
        layout.addWidget(file_group)
        
        # جدول الجداول المستخرجة
        tables_group = QGroupBox('الجداول المكتشفة')
        tables_layout = QVBoxLayout(tables_group)
        
        self.tables_list = QListWidget()
        self.tables_list.itemClicked.connect(self.on_table_selected)
        tables_layout.addWidget(self.tables_list)
        
        layout.addWidget(tables_group)
        
        # معاينة الجدول المحدد
        preview_group = QGroupBox('محتوى الجدول')
        preview_layout = QVBoxLayout(preview_group)
        
        self.table_preview = QTableWidget()
        self.table_preview.setLayoutDirection(Qt.RightToLeft)
        preview_layout.addWidget(self.table_preview)
        
        # أزرار الاستيراد
        import_layout = QHBoxLayout()
        
        self.cmb_import_category = QComboBox()
        self.cmb_import_category.addItems(self.DEFAULT_CATEGORIES)
        import_layout.addWidget(QLabel('التصنيف:'))
        import_layout.addWidget(self.cmb_import_category)
        
        btn_import = QPushButton('📥 استيراد الجدول المحدد')
        btn_import.clicked.connect(self.import_selected_table)
        import_layout.addWidget(btn_import)
        
        btn_import_all = QPushButton('📥 استيراد الكل (تلقائي)')
        btn_import_all.clicked.connect(self.import_all_tables)
        import_layout.addWidget(btn_import_all)
        
        preview_layout.addLayout(import_layout)
        layout.addWidget(preview_group)
        
        return widget
    
    def create_manual_tab(self):
        """تبويب الإضافة اليدوية"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        
        # القائمة الجانبية - التصنيفات والنماذج
        sidebar = QWidget()
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar.setMaximumWidth(350)
        
        # التصنيفات
        cat_group = QGroupBox('التصنيفات')
        cat_layout = QVBoxLayout(cat_group)
        
        self.categories_list = QListWidget()
        self.categories_list.addItems(self.DEFAULT_CATEGORIES)
        self.categories_list.currentRowChanged.connect(self.on_category_changed)
        cat_layout.addWidget(self.categories_list)
        
        cat_buttons = QHBoxLayout()
        btn_add_cat = QPushButton('➕')
        btn_add_cat.setToolTip('إضافة تصنيف')
        btn_add_cat.clicked.connect(self.add_category)
        cat_buttons.addWidget(btn_add_cat)
        
        btn_del_cat = QPushButton('➖')
        btn_del_cat.setToolTip('حذف تصنيف')
        btn_del_cat.clicked.connect(self.delete_category)
        cat_buttons.addWidget(btn_del_cat)
        cat_layout.addLayout(cat_buttons)
        
        sidebar_layout.addWidget(cat_group)
        
        # النماذج في التصنيف
        templates_group = QGroupBox('النماذج')
        templates_layout = QVBoxLayout(templates_group)
        
        self.templates_list = QListWidget()
        self.templates_list.currentRowChanged.connect(self.on_template_selected)
        self.templates_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.templates_list.customContextMenuRequested.connect(self.show_template_context_menu)
        templates_layout.addWidget(self.templates_list)
        
        tmpl_buttons = QHBoxLayout()
        btn_add_tmpl = QPushButton('➕ إضافة نموذج')
        btn_add_tmpl.clicked.connect(self.add_template)
        tmpl_buttons.addWidget(btn_add_tmpl)
        
        btn_del_tmpl = QPushButton('🗑️ حذف')
        btn_del_tmpl.clicked.connect(self.delete_template)
        tmpl_buttons.addWidget(btn_del_tmpl)
        templates_layout.addLayout(tmpl_buttons)
        
        sidebar_layout.addWidget(templates_group)
        layout.addWidget(sidebar)
        
        # منطقة التحرير
        editor = QWidget()
        editor_layout = QVBoxLayout(editor)
        
        form_group = QGroupBox('تحرير النموذج')
        form_layout = QFormLayout(form_group)
        
        self.txt_num = QLineEdit()
        self.txt_num.setPlaceholderText('مثال: 1، 2، 30')
        form_layout.addRow('الرقم:', self.txt_num)
        
        self.txt_keyword = QLineEdit()
        self.txt_keyword.setPlaceholderText('مثال: لدي1، قائم، أدعى')
        form_layout.addRow('الكلمة المفتاحية:', self.txt_keyword)
        
        self.txt_content = QTextEdit()
        self.txt_content.setPlaceholderText('محتوى النموذج...')
        self.txt_content.setMinimumHeight(300)
        form_layout.addRow('المحتوى:', self.txt_content)
        
        editor_layout.addWidget(form_group)
        
        # أزرار الحفظ
        save_layout = QHBoxLayout()
        btn_save_template = QPushButton('💾 حفظ التعديلات')
        btn_save_template.clicked.connect(self.save_current_template)
        save_layout.addWidget(btn_save_template)
        
        btn_clear = QPushButton('🔄 مسح الحقول')
        btn_clear.clicked.connect(self.clear_editor)
        save_layout.addWidget(btn_clear)
        
        editor_layout.addLayout(save_layout)
        layout.addWidget(editor)
        
        return widget
    
    def create_preview_tab(self):
        """تبويب المعاينة"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # إحصائيات
        stats_group = QGroupBox('📊 إحصائيات')
        stats_layout = QHBoxLayout(stats_group)
        
        self.lbl_total_templates = QLabel('إجمالي النماذج: 0')
        stats_layout.addWidget(self.lbl_total_templates)
        
        self.lbl_total_categories = QLabel('التصنيفات: 0')
        stats_layout.addWidget(self.lbl_total_categories)
        
        layout.addWidget(stats_group)
        
        # معاينة الكود
        preview_group = QGroupBox('معاينة الكود')
        preview_layout = QVBoxLayout(preview_group)
        
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel('الصيغة:'))
        
        self.cmb_format = QComboBox()
        self.cmb_format.addItems(['JavaScript (templatesData)', 'JSON'])
        self.cmb_format.currentIndexChanged.connect(self.update_preview)
        format_layout.addWidget(self.cmb_format)
        
        btn_refresh = QPushButton('🔄 تحديث المعاينة')
        btn_refresh.clicked.connect(self.update_preview)
        format_layout.addWidget(btn_refresh)
        
        format_layout.addStretch()
        preview_layout.addLayout(format_layout)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont('Courier New', 10))
        self.preview_text.setLayoutDirection(Qt.LeftToRight)
        preview_layout.addWidget(self.preview_text)
        
        # أزرار النسخ
        copy_layout = QHBoxLayout()
        btn_copy = QPushButton('📋 نسخ للحافظة')
        btn_copy.clicked.connect(self.copy_to_clipboard)
        copy_layout.addWidget(btn_copy)
        preview_layout.addLayout(copy_layout)
        
        layout.addWidget(preview_group)
        
        return widget
    
    def create_export_tab(self):
        """تبويب التصدير"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # خيارات التصدير
        options_group = QGroupBox('خيارات التصدير')
        options_layout = QFormLayout(options_group)
        
        self.cmb_export_format = QComboBox()
        self.cmb_export_format.addItems([
            'JavaScript (.js) - للاستخدام في HTML',
            'JSON (.json) - بيانات خام'
        ])
        options_layout.addRow('الصيغة:', self.cmb_export_format)
        
        self.chk_minify = QCheckBox('ضغط الكود (minify)')
        options_layout.addRow('', self.chk_minify)
        
        layout.addWidget(options_group)
        
        # أزرار التصدير
        export_group = QGroupBox('تصدير')
        export_layout = QVBoxLayout(export_group)
        
        btn_export_file = QPushButton('💾 حفظ كملف')
        btn_export_file.clicked.connect(self.export_to_file)
        btn_export_file.setMinimumHeight(50)
        export_layout.addWidget(btn_export_file)
        
        btn_export_clipboard = QPushButton('📋 نسخ للحافظة')
        btn_export_clipboard.clicked.connect(self.copy_to_clipboard)
        btn_export_clipboard.setMinimumHeight(50)
        export_layout.addWidget(btn_export_clipboard)
        
        layout.addWidget(export_group)
        layout.addStretch()
        
        return widget
    
    # ==================== وظائف الاستيراد ====================
    
    def open_word_file(self):
        """فتح ملف وورد"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'اختر ملف وورد', '', 'Word Files (*.docx *.doc);;All Files (*)'
        )
        if file_path:
            self.current_file = file_path
            self.lbl_file.setText(os.path.basename(file_path))
            
            # إذا كان الملف .doc (قديم) نحوله إلى .docx
            if file_path.lower().endswith('.doc') and not file_path.lower().endswith('.docx'):
                converted_path = self.convert_doc_to_docx(file_path)
                if converted_path:
                    self.extract_tables_from_word(converted_path)
                else:
                    QMessageBox.critical(self, 'خطأ', 
                        'لم يتم التحويل. جرّب:\n'
                        '1. تحويل الملف يدوياً إلى .docx من Word\n'
                        '2. أو تثبيت LibreOffice')
            else:
                self.extract_tables_from_word(file_path)
    
    def convert_doc_to_docx(self, doc_path):
        """تحويل ملف .doc إلى .docx"""
        import subprocess
        import tempfile
        
        # المسار المؤقت للملف المحول
        temp_dir = tempfile.gettempdir()
        docx_path = os.path.join(temp_dir, os.path.basename(doc_path) + 'x')
        
        # محاولة 1: استخدام LibreOffice
        try:
            # البحث عن LibreOffice
            libreoffice_paths = [
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
                '/usr/bin/libreoffice',
                '/usr/bin/soffice',
                'libreoffice',
                'soffice'
            ]
            
            soffice = None
            for path in libreoffice_paths:
                if os.path.exists(path) or self.command_exists(path):
                    soffice = path
                    break
            
            if soffice:
                self.status_bar.showMessage('جاري تحويل الملف...', 0)
                QApplication.processEvents()
                
                result = subprocess.run([
                    soffice,
                    '--headless',
                    '--convert-to', 'docx',
                    '--outdir', temp_dir,
                    doc_path
                ], capture_output=True, timeout=60)
                
                if os.path.exists(docx_path):
                    self.status_bar.showMessage('تم تحويل الملف بنجاح', 3000)
                    return docx_path
        except Exception as e:
            print(f"LibreOffice error: {e}")
        
        # محاولة 2: استخدام Word COM (Windows فقط)
        if sys.platform == 'win32':
            try:
                import win32com.client
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(os.path.abspath(doc_path))
                doc.SaveAs2(docx_path, FileFormat=16)  # 16 = docx
                doc.Close()
                word.Quit()
                
                if os.path.exists(docx_path):
                    self.status_bar.showMessage('تم تحويل الملف بنجاح', 3000)
                    return docx_path
            except ImportError:
                QMessageBox.warning(self, 'تنبيه', 
                    'لتحويل ملفات .doc، ثبّت:\n'
                    'pip install pywin32\n\n'
                    'أو استخدم LibreOffice')
            except Exception as e:
                print(f"Word COM error: {e}")
        
        return None
    
    def command_exists(self, cmd):
        """التحقق من وجود أمر"""
        import shutil
        return shutil.which(cmd) is not None
    
    def extract_tables_from_word(self, file_path):
        """استخراج الجداول من ملف وورد"""
        try:
            doc = Document(file_path)
            self.word_tables = []
            self.tables_list.clear()
            
            for i, table in enumerate(doc.tables):
                rows = []
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    rows.append(cells)
                
                if rows:
                    self.word_tables.append(rows)
                    # تحديد اسم الجدول من أول خلية
                    first_text = rows[0][0] if rows[0] else f'جدول {i+1}'
                    preview = first_text[:50] + '...' if len(first_text) > 50 else first_text
                    self.tables_list.addItem(f'جدول {i+1}: {preview} ({len(rows)} صف)')
            
            self.status_bar.showMessage(f'تم استخراج {len(self.word_tables)} جدول', 5000)
            
        except Exception as e:
            QMessageBox.critical(self, 'خطأ', f'فشل في قراءة الملف:\n{str(e)}')
    
    def on_table_selected(self, item):
        """عند اختيار جدول"""
        idx = self.tables_list.currentRow()
        if idx >= 0 and idx < len(self.word_tables):
            self.show_table_preview(self.word_tables[idx])
    
    def show_table_preview(self, table_data):
        """عرض معاينة الجدول"""
        self.table_preview.clear()
        if not table_data:
            return
        
        # تحديد عدد الأعمدة
        max_cols = max(len(row) for row in table_data)
        self.table_preview.setRowCount(len(table_data))
        self.table_preview.setColumnCount(max_cols)
        
        for i, row in enumerate(table_data):
            for j, cell in enumerate(row):
                item = QTableWidgetItem(cell[:100])  # اقتصار النص
                self.table_preview.setItem(i, j, item)
        
        self.table_preview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
    def import_selected_table(self):
        """استيراد الجدول المحدد"""
        idx = self.tables_list.currentRow()
        if idx < 0:
            QMessageBox.warning(self, 'تنبيه', 'اختر جدولاً أولاً')
            return
        
        category = self.cmb_import_category.currentText()
        table_data = self.word_tables[idx]
        
        imported = 0
        for row in table_data[1:]:  # تخطي الصف الأول (العنوان غالباً)
            if len(row) >= 3:
                num = row[0].strip()
                keyword = row[1].strip()
                content = row[2].strip()
                
                if content and len(content) > 10:
                    template = Template(num, keyword, content, category)
                    self.templates[category].append(template)
                    imported += 1
        
        self.update_templates_list()
        self.update_status()
        QMessageBox.information(self, 'تم', f'تم استيراد {imported} نموذج إلى "{category}"')
    
    def import_all_tables(self):
        """استيراد كل الجداول تلقائياً"""
        if not hasattr(self, 'word_tables') or not self.word_tables:
            QMessageBox.warning(self, 'تنبيه', 'لا توجد جداول للاستيراد')
            return
        
        # قاموس لربط أسماء الجداول بالتصنيفات
        category_mapping = {
            'الدعوى': 'الدعوى',
            'صندوق الدعوى': 'الدعوى',
            'الإجابة': 'الإجابة',
            'صندوق الإجابة': 'الإجابة',
            'المرافعة': 'المرافعة',
            'صندوق المرافعة': 'المرافعة',
            'الأسباب': 'الأسباب',
            'صندوق الأسباب': 'الأسباب',
            'الحكم': 'الحكم',
            'صندوق الحكم': 'الحكم',
            'الشهادة': 'الشهادة',
            'الصلح': 'الصلح',
            'اليمين': 'اليمين',
            'النكول': 'النكول',
            'الكفالة': 'الكفالة',
            'الالتماس': 'الالتماس',
            'الشطب': 'الشطب',
            'الغياب': 'الغياب',
            'الاختصاص': 'الاختصاص',
            'التمويل': 'التمويل',
            'العقارات': 'العقارات',
            'المشاكل': 'المشاكل_التقنية',
        }
        
        total_imported = 0
        
        for table_data in self.word_tables:
            if not table_data:
                continue
            
            # محاولة تحديد التصنيف من أول صف
            first_cell = table_data[0][0] if table_data[0] else ''
            category = 'الدعوى'  # افتراضي
            
            for key, cat in category_mapping.items():
                if key in first_cell:
                    category = cat
                    break
            
            # استيراد الصفوف
            for row in table_data[1:]:
                if len(row) >= 3:
                    num = row[0].strip()
                    keyword = row[1].strip()
                    content = row[2].strip()
                    
                    if content and len(content) > 10:
                        template = Template(num, keyword, content, category)
                        self.templates[category].append(template)
                        total_imported += 1
        
        self.update_templates_list()
        self.update_status()
        QMessageBox.information(self, 'تم', f'تم استيراد {total_imported} نموذج')
    
    # ==================== وظائف التحرير اليدوي ====================
    
    def on_category_changed(self, index):
        """عند تغيير التصنيف"""
        self.update_templates_list()
    
    def update_templates_list(self):
        """تحديث قائمة النماذج"""
        self.templates_list.clear()
        
        current_row = self.categories_list.currentRow()
        if current_row < 0:
            return
        
        category = self.categories_list.item(current_row).text()
        
        if category in self.templates:
            for i, tmpl in enumerate(self.templates[category]):
                display = f"[{tmpl.num}] {tmpl.keyword}: {tmpl.content[:40]}..."
                self.templates_list.addItem(display)
    
    def on_template_selected(self, index):
        """عند اختيار نموذج"""
        if index < 0:
            return
        
        current_cat = self.categories_list.currentItem()
        if not current_cat:
            return
        
        category = current_cat.text()
        
        if category in self.templates and index < len(self.templates[category]):
            tmpl = self.templates[category][index]
            self.txt_num.setText(tmpl.num)
            self.txt_keyword.setText(tmpl.keyword)
            self.txt_content.setPlainText(tmpl.content)
    
    def add_category(self):
        """إضافة تصنيف جديد"""
        name, ok = QInputDialog.getText(self, 'تصنيف جديد', 'اسم التصنيف:')
        if ok and name:
            if name not in self.templates:
                self.templates[name] = []
                self.categories_list.addItem(name)
                self.cmb_import_category.addItem(name)
    
    def delete_category(self):
        """حذف التصنيف المحدد"""
        current = self.categories_list.currentItem()
        if not current:
            return
        
        name = current.text()
        reply = QMessageBox.question(
            self, 'تأكيد الحذف',
            f'هل تريد حذف التصنيف "{name}" وجميع نماذجه؟',
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if name in self.templates:
                del self.templates[name]
            self.categories_list.takeItem(self.categories_list.currentRow())
            
            # حذف من القائمة المنسدلة
            idx = self.cmb_import_category.findText(name)
            if idx >= 0:
                self.cmb_import_category.removeItem(idx)
    
    def add_template(self):
        """إضافة نموذج جديد"""
        current_cat = self.categories_list.currentItem()
        if not current_cat:
            QMessageBox.warning(self, 'تنبيه', 'اختر تصنيفاً أولاً')
            return
        
        category = current_cat.text()
        
        # إنشاء نموذج فارغ
        tmpl = Template('', '', '', category)
        self.templates[category].append(tmpl)
        self.update_templates_list()
        
        # تحديد النموذج الجديد
        self.templates_list.setCurrentRow(len(self.templates[category]) - 1)
        self.clear_editor()
    
    def delete_template(self):
        """حذف النموذج المحدد"""
        current_cat = self.categories_list.currentItem()
        current_tmpl = self.templates_list.currentRow()
        
        if not current_cat or current_tmpl < 0:
            return
        
        category = current_cat.text()
        
        reply = QMessageBox.question(
            self, 'تأكيد الحذف',
            'هل تريد حذف هذا النموذج؟',
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            if category in self.templates:
                del self.templates[category][current_tmpl]
                self.update_templates_list()
                self.clear_editor()
                self.update_status()
    
    def save_current_template(self):
        """حفظ التعديلات على النموذج الحالي"""
        current_cat = self.categories_list.currentItem()
        current_idx = self.templates_list.currentRow()
        
        if not current_cat or current_idx < 0:
            QMessageBox.warning(self, 'تنبيه', 'اختر نموذجاً أولاً')
            return
        
        category = current_cat.text()
        
        if category in self.templates and current_idx < len(self.templates[category]):
            tmpl = self.templates[category][current_idx]
            tmpl.num = self.txt_num.text().strip()
            tmpl.keyword = self.txt_keyword.text().strip()
            tmpl.content = self.txt_content.toPlainText().strip()
            
            self.update_templates_list()
            self.templates_list.setCurrentRow(current_idx)
            self.update_status()
            self.status_bar.showMessage('تم حفظ التعديلات', 3000)
    
    def clear_editor(self):
        """مسح حقول التحرير"""
        self.txt_num.clear()
        self.txt_keyword.clear()
        self.txt_content.clear()
    
    def show_template_context_menu(self, pos):
        """قائمة السياق للنماذج"""
        menu = QMenu(self)
        
        action_copy = menu.addAction('📋 نسخ المحتوى')
        action_copy.triggered.connect(self.copy_template_content)
        
        action_delete = menu.addAction('🗑️ حذف')
        action_delete.triggered.connect(self.delete_template)
        
        menu.exec_(self.templates_list.mapToGlobal(pos))
    
    def copy_template_content(self):
        """نسخ محتوى النموذج"""
        content = self.txt_content.toPlainText()
        if content:
            QApplication.clipboard().setText(content)
            self.status_bar.showMessage('تم النسخ للحافظة', 3000)
    
    # ==================== وظائف المعاينة والتصدير ====================
    
    def update_preview(self):
        """تحديث المعاينة"""
        format_type = self.cmb_format.currentIndex()
        
        # حساب الإحصائيات
        total = sum(len(tmpls) for tmpls in self.templates.values())
        non_empty_cats = sum(1 for tmpls in self.templates.values() if tmpls)
        
        self.lbl_total_templates.setText(f'إجمالي النماذج: {total}')
        self.lbl_total_categories.setText(f'التصنيفات: {non_empty_cats}')
        
        # إنشاء الكود
        if format_type == 0:  # JavaScript
            code = self.generate_js_code()
        else:  # JSON
            code = self.generate_json_code()
        
        self.preview_text.setPlainText(code)
    
    def generate_js_code(self, minify=False):
        """إنشاء كود JavaScript"""
        data = {}
        for cat, tmpls in self.templates.items():
            if tmpls:
                data[cat] = [t.to_dict() for t in tmpls]
        
        if minify:
            json_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
        else:
            json_str = json.dumps(data, ensure_ascii=False, indent=4)
        
        return f"const templatesData = {json_str};"
    
    def generate_json_code(self, minify=False):
        """إنشاء كود JSON"""
        data = {}
        for cat, tmpls in self.templates.items():
            if tmpls:
                data[cat] = [t.to_dict() for t in tmpls]
        
        if minify:
            return json.dumps(data, ensure_ascii=False, separators=(',', ':'))
        else:
            return json.dumps(data, ensure_ascii=False, indent=4)
    
    def copy_to_clipboard(self):
        """نسخ للحافظة"""
        format_type = self.cmb_format.currentIndex() if hasattr(self, 'cmb_format') else 0
        minify = self.chk_minify.isChecked() if hasattr(self, 'chk_minify') else False
        
        if format_type == 0:
            code = self.generate_js_code(minify)
        else:
            code = self.generate_json_code(minify)
        
        QApplication.clipboard().setText(code)
        self.status_bar.showMessage('تم النسخ للحافظة ✓', 3000)
        QMessageBox.information(self, 'تم', 'تم نسخ الكود للحافظة بنجاح')
    
    def export_to_file(self):
        """تصدير كملف"""
        format_idx = self.cmb_export_format.currentIndex()
        minify = self.chk_minify.isChecked()
        
        if format_idx == 0:  # JavaScript
            ext = 'js'
            code = self.generate_js_code(minify)
        else:  # JSON
            ext = 'json'
            code = self.generate_json_code(minify)
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, 'حفظ الملف', f'templatesData.{ext}',
            f'{ext.upper()} Files (*.{ext});;All Files (*)'
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(code)
                QMessageBox.information(self, 'تم', f'تم حفظ الملف:\n{file_path}')
            except Exception as e:
                QMessageBox.critical(self, 'خطأ', f'فشل في الحفظ:\n{str(e)}')
    
    # ==================== حفظ/تحميل المشروع ====================
    
    def save_project(self):
        """حفظ المشروع"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, 'حفظ المشروع', 'templates_project.json',
            'JSON Files (*.json);;All Files (*)'
        )
        
        if file_path:
            try:
                data = {
                    'version': '1.0',
                    'templates': {cat: [t.to_dict() for t in tmpls] for cat, tmpls in self.templates.items()}
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.status_bar.showMessage('تم حفظ المشروع', 3000)
            except Exception as e:
                QMessageBox.critical(self, 'خطأ', f'فشل في الحفظ:\n{str(e)}')
    
    def load_project(self):
        """تحميل مشروع"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 'تحميل مشروع', '', 'JSON Files (*.json);;All Files (*)'
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.templates = {}
                for cat, tmpls in data.get('templates', {}).items():
                    self.templates[cat] = [
                        Template(t.get('num', ''), t.get('keyword', ''), t.get('content', ''), cat)
                        for t in tmpls
                    ]
                
                # تحديث القوائم
                self.categories_list.clear()
                self.categories_list.addItems(list(self.templates.keys()))
                
                self.cmb_import_category.clear()
                self.cmb_import_category.addItems(list(self.templates.keys()))
                
                self.update_status()
                self.status_bar.showMessage('تم تحميل المشروع', 3000)
                
            except Exception as e:
                QMessageBox.critical(self, 'خطأ', f'فشل في التحميل:\n{str(e)}')
    
    # ==================== المظهر ====================
    
    def toggle_theme(self):
        """تبديل الوضع الليلي"""
        self.is_dark_mode = not self.is_dark_mode
        if self.is_dark_mode:
            self.apply_dark_theme()
            self.btn_theme.setText('☀️ الوضع النهاري')
        else:
            self.apply_light_theme()
            self.btn_theme.setText('🌙 الوضع الليلي')
    
    def apply_light_theme(self):
        """تطبيق المظهر الفاتح"""
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #f8f6f1;
                color: #2c2c2c;
                font-family: 'Tajawal', 'Segoe UI', sans-serif;
                font-size: 14px;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e0ddd5;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top right;
                padding: 0 10px;
                color: #1a5f4a;
            }
            QPushButton {
                background-color: #1a5f4a;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2d8b6e;
            }
            QPushButton:pressed {
                background-color: #0d3d2f;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox {
                border: 2px solid #e0ddd5;
                border-radius: 6px;
                padding: 8px;
                background-color: #ffffff;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border-color: #1a5f4a;
            }
            QListWidget, QTableWidget {
                border: 2px solid #e0ddd5;
                border-radius: 6px;
                background-color: #ffffff;
            }
            QListWidget::item:selected, QTableWidget::item:selected {
                background-color: #1a5f4a;
                color: white;
            }
            QTabWidget::pane {
                border: 2px solid #e0ddd5;
                border-radius: 8px;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #e0ddd5;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background-color: #1a5f4a;
                color: white;
            }
            QStatusBar {
                background-color: #1a5f4a;
                color: white;
            }
            QHeaderView::section {
                background-color: #1a5f4a;
                color: white;
                padding: 8px;
                border: none;
            }
        """)
    
    def apply_dark_theme(self):
        """تطبيق المظهر الداكن"""
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1a1a1a;
                color: #f0f0f0;
                font-family: 'Tajawal', 'Segoe UI', sans-serif;
                font-size: 14px;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #404040;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #2d2d2d;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top right;
                padding: 0 10px;
                color: #2d8b6e;
            }
            QPushButton {
                background-color: #2d8b6e;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3da37f;
            }
            QPushButton:pressed {
                background-color: #1a5f4a;
            }
            QLineEdit, QTextEdit, QComboBox, QSpinBox {
                border: 2px solid #404040;
                border-radius: 6px;
                padding: 8px;
                background-color: #2d2d2d;
                color: #f0f0f0;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border-color: #2d8b6e;
            }
            QListWidget, QTableWidget {
                border: 2px solid #404040;
                border-radius: 6px;
                background-color: #2d2d2d;
                color: #f0f0f0;
            }
            QListWidget::item:selected, QTableWidget::item:selected {
                background-color: #2d8b6e;
                color: white;
            }
            QTabWidget::pane {
                border: 2px solid #404040;
                border-radius: 8px;
                background-color: #2d2d2d;
            }
            QTabBar::tab {
                background-color: #404040;
                color: #f0f0f0;
                padding: 10px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
            }
            QTabBar::tab:selected {
                background-color: #2d8b6e;
                color: white;
            }
            QStatusBar {
                background-color: #2d8b6e;
                color: white;
            }
            QHeaderView::section {
                background-color: #2d8b6e;
                color: white;
                padding: 8px;
                border: none;
            }
            QComboBox QAbstractItemView {
                background-color: #2d2d2d;
                color: #f0f0f0;
                selection-background-color: #2d8b6e;
            }
        """)
    
    def update_status(self):
        """تحديث شريط الحالة"""
        total = sum(len(tmpls) for tmpls in self.templates.values())
        self.status_bar.showMessage(f'إجمالي النماذج: {total}')


def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    
    # تعيين الخط
    font = QFont('Tajawal', 11)
    app.setFont(font)
    
    window = TemplateConverter()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
