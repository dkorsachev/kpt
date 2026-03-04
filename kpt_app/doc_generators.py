import re
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class DocumentGenerator:
    """Класс для генерации документов на основе шаблонов"""
    
    def __init__(self):
        self.markers = {
            'kadastr_number': '##KADASTR_NUMBER##',
            'contract_number': '##CONTRACT_NUMBER##',
            'contract_date': '##CONTRACT_DATE##',
            'contract_date_short': '##CONTRACT_DATE_SHORT##',
            'current_date': '##CURRENT_DATE##',
            'current_date_short': '##CURRENT_DATE_SHORT##',
            'table_place': '##TABLE_PLACE##',
        }
    
    def format_date_for_doc(self, date_obj):
        """Форматирование даты для документа (полный формат)"""
        if isinstance(date_obj, str):
            return date_obj
        months = [
            'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
            'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
        ]
        return f"{date_obj.day} {months[date_obj.month-1]} {date_obj.year} года"
    
    def format_date_short(self, date_obj):
        """Форматирование даты в короткий формат ДД.ММ.ГГГГ"""
        if isinstance(date_obj, str):
            return date_obj
        return date_obj.strftime('%d.%m.%Y')
    
    def replace_in_paragraph_preserve_format(self, paragraph, replacements):
        """Замена текста в параграфе с сохранением оригинального форматирования"""
        full_text = paragraph.text
        new_text = full_text
        
        for key, value in replacements.items():
            if key in new_text:
                new_text = new_text.replace(key, value)
        
        if full_text != new_text and new_text.strip():
            # Сохраняем форматирование первого run
            font_name = 'Times New Roman'
            font_size = Pt(12)
            bold = False
            italic = False
            underline = False
            
            if paragraph.runs:
                original_run = paragraph.runs[0]
                if original_run.font.name:
                    font_name = original_run.font.name
                if original_run.font.size:
                    font_size = original_run.font.size
                bold = original_run.font.bold or False
                italic = original_run.font.italic or False
                underline = original_run.font.underline or False
            
            # Очищаем все runs
            paragraph.clear()
            
            # Добавляем новый run с сохраненным форматированием
            run = paragraph.add_run(new_text)
            run.font.name = font_name
            run.font.size = font_size
            run.font.bold = bold
            run.font.italic = italic
            run.font.underline = underline
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    def set_cell_font_size_10(self, cell, bold=False):
        """Установка шрифта 10 для ячейки таблицы (для служебной записки)"""
        for paragraph in cell.paragraphs:
            text = paragraph.text
            paragraph.clear()
            run = paragraph.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            run.font.bold = bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
            
            # Компактное форматирование
            paragraph.paragraph_format.line_spacing = Pt(12)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
    
    def create_table_at_marker_sluzhebka(self, doc, marker_paragraph, rows_data):
        """Создание таблицы для служебной записки (шрифт 10)"""
        parent = marker_paragraph._element.getparent()
        
        # Создаем таблицу
        table = doc.add_table(rows=1 + len(rows_data), cols=4)
        table.style = 'Table Grid'
        table.autofit = False
        
        # Выравнивание по центру
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is not None:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            tblPr.append(jc)
        
        # Ширина колонок
        if len(table.columns) >= 4:
            table.columns[0].width = Cm(1.2)
            table.columns[1].width = Cm(3.5)
            table.columns[2].width = Cm(4.5)
            table.columns[3].width = Cm(5.5)
        
        # Заголовки таблицы (шрифт 10, жирный)
        headers = ['№ п/п', 'Тип документа', 'Кадастровый номер', 'Основание']
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            self.set_cell_font_size_10(cell, bold=True)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Данные таблицы (шрифт 10)
        for i, row in enumerate(rows_data, 1):
            cad_num = row['cadastral_number']
            
            # Определение типа документа по формату кадастрового номера
            if re.match(r'^\d{2}:\d{2}:\d{7}$', cad_num):
                # Формат 95:ХХ:ХХХХХХХ - кадастровый квартал
                doc_type = "Кадастровый квартал"
            elif re.match(r'^\d{2}:\d{2}:\d{7}:\d{1,5}$', cad_num):
                # Формат 95:ХХ:ХХХХХХХ:ХХХХ - земельный участок
                doc_type = "Выписка ЕГРН"
            else:
                # На всякий случай
                doc_type = "Объект недвижимости"
            
            contract_date = datetime.strptime(row['contract_date'], '%Y-%m-%d')
            date_str = self.format_date_for_doc(contract_date)
            foundation = f"{row['contract_number']} от {date_str}"
            
            row_cells = table.rows[i].cells
            row_cells[0].text = str(i)
            row_cells[1].text = doc_type
            row_cells[2].text = cad_num
            row_cells[3].text = foundation
            
            for cell in row_cells:
                self.set_cell_font_size_10(cell)
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Вставляем таблицу после маркера
        marker_paragraph._element.addnext(table._element)
        
        # Удаляем параграф с маркером
        parent.remove(marker_paragraph._element)
    
    def generate_sluzhebka(self, template_path, rows_data, current_date):
        """Генерация служебной записки"""
        doc = Document(template_path)
        
        # Общие замены (даты) - с сохранением оригинального форматирования
        replacements = {
            self.markers['current_date']: self.format_date_for_doc(current_date),
            self.markers['current_date_short']: self.format_date_short(current_date),
        }
        
        # Сначала заменяем все маркеры дат с сохранением форматирования
        for paragraph in doc.paragraphs:
            if self.markers['table_place'] not in paragraph.text and '##TABLE_ROW##' not in paragraph.text:
                self.replace_in_paragraph_preserve_format(paragraph, replacements)
        
        # Ищем и обрабатываем маркер таблицы
        marker_paragraphs = []
        for paragraph in doc.paragraphs:
            if self.markers['table_place'] in paragraph.text or '##TABLE_ROW##' in paragraph.text:
                marker_paragraphs.append(paragraph)
        
        # Создаем таблицы на месте каждого маркера
        for marker_paragraph in marker_paragraphs:
            self.create_table_at_marker_sluzhebka(doc, marker_paragraph, rows_data)
        
        return doc
    
    def generate_kpt(self, template_path, cadastral_number, contract_number, contract_date, current_date):
        """Генерация запроса на кадастровый квартал с сохранением форматирования шаблона"""
        doc = Document(template_path)
        
        replacements = {
            self.markers['kadastr_number']: cadastral_number,
            self.markers['contract_number']: contract_number,
            self.markers['contract_date']: self.format_date_for_doc(contract_date),
            self.markers['contract_date_short']: self.format_date_short(contract_date),
            self.markers['current_date']: self.format_date_for_doc(current_date),
            self.markers['current_date_short']: self.format_date_short(current_date),
        }
        
        # Заменяем с сохранением оригинального форматирования
        for paragraph in doc.paragraphs:
            self.replace_in_paragraph_preserve_format(paragraph, replacements)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_in_paragraph_preserve_format(paragraph, replacements)
        
        return doc
    
    def generate_zu(self, template_path, cadastral_number, contract_number, contract_date, current_date):
        """Генерация запроса на земельный участок с сохранением форматирования шаблона"""
        doc = Document(template_path)
        
        replacements = {
            self.markers['kadastr_number']: cadastral_number,
            self.markers['contract_number']: contract_number,
            self.markers['contract_date']: self.format_date_for_doc(contract_date),
            self.markers['contract_date_short']: self.format_date_short(contract_date),
            self.markers['current_date']: self.format_date_for_doc(current_date),
            self.markers['current_date_short']: self.format_date_short(current_date),
        }
        
        # Заменяем с сохранением оригинального форматирования
        for paragraph in doc.paragraphs:
            self.replace_in_paragraph_preserve_format(paragraph, replacements)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_in_paragraph_preserve_format(paragraph, replacements)
        
        return doc
    
    def convert_to_pdf(self, docx_path, pdf_path):
        """Конвертация DOCX в PDF"""
        try:
            # Пробуем использовать docx2pdf
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            return True
        except ImportError:
            try:
                # Альтернативный способ для Windows
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
                word.Quit()
                return True
            except:
                # Если ничего не работает, возвращаем False
                return False