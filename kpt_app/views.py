import os
import re
import zipfile
import json
from io import BytesIO
from datetime import datetime
from django.shortcuts import render
from django.http import FileResponse
from django.conf import settings
from .forms import CadastralNumberForm
from .doc_generators import DocumentGenerator
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading


# Создаем один генератор для всего запроса (переиспользование)
generator = DocumentGenerator()


def index(request):
    form = CadastralNumberForm()
    return render(request, 'kpt_app/index.html', {'form': form})


def generate_documents(request):
    if request.method != 'POST':
        form = CadastralNumberForm()
        return render(request, 'kpt_app/index.html', {'form': form})
    
    form = CadastralNumberForm(request.POST)
    if not form.is_valid():
        return render(request, 'kpt_app/index.html', {'form': form})
    
    # Получаем данные
    rows_data = form.cleaned_data['rows_json']
    current_date = datetime.now()
    
    # Пути к шаблонам
    templates_dir = os.path.join(settings.BASE_DIR, 'templates_docs')
    
    # Кэшируем пути
    sluzhebka_template = os.path.join(templates_dir, 'sluzhebka_template.docx')
    if not os.path.exists(sluzhebka_template):
        sluzhebka_template = os.path.join(templates_dir, 'sluzhebka_template.docx.docx')
    
    kpt_template = os.path.join(templates_dir, 'kpt_template.docx')
    if not os.path.exists(kpt_template):
        kpt_template = os.path.join(templates_dir, 'kpt_template.docx.docx')
    
    zu_template = os.path.join(templates_dir, 'zu_template.docx')
    if not os.path.exists(zu_template):
        zu_template = os.path.join(templates_dir, 'zu_template.docx.docx')
    
    if not os.path.exists(sluzhebka_template):
        return render(request, 'kpt_app/index.html', {
            'form': form,
            'error': f'Файл шаблона не найден: sluzhebka_template.docx'
        })
    
    # Создаем временную директорию
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Генерируем служебную записку
            try:
                sluzhebka_doc = generator.generate_sluzhebka(
                    sluzhebka_template,
                    rows_data,
                    current_date
                )
                sluzhebka_docx = os.path.join(temp_dir, "Запрос из ЕГРН.docx")
                sluzhebka_doc.save(sluzhebka_docx)
                zip_file.write(sluzhebka_docx, "Запрос из ЕГРН.docx")
                
                # Конвертируем в PDF в отдельном потоке
                sluzhebka_pdf = os.path.join(temp_dir, "Запрос из ЕГРН.pdf")
                if generator.convert_to_pdf(sluzhebka_docx, sluzhebka_pdf):
                    zip_file.write(sluzhebka_pdf, "Запрос из ЕГРН.pdf")
            except Exception as e:
                return render(request, 'kpt_app/index.html', {
                    'form': form,
                    'error': f'Ошибка при генерации служебной записки: {str(e)}'
                })
            
            # Параллельная генерация файлов для каждой строки
            def generate_row_file(row):
                try:
                    cad_num = row['cadastral_number']
                    contract_number = row['contract_number']
                    contract_date = datetime.strptime(row['contract_date'], '%Y-%m-%d')
                    
                    if re.match(r'^\d{2}:\d{2}:\d{7}$', cad_num):
                        template = kpt_template
                        doc = generator.generate_kpt(
                            template, cad_num, contract_number, contract_date, current_date
                        )
                        num_parts = cad_num.replace(':', '_')
                        filename = f"Запрос_на_кад_квартал_{num_parts}.docx"
                    else:
                        template = zu_template
                        doc = generator.generate_zu(
                            template, cad_num, contract_number, contract_date, current_date
                        )
                        num_parts = cad_num.replace(':', '_')
                        filename = f"Запрос_выписки_ЗУ_{num_parts}.docx"
                    
                    if not os.path.exists(template):
                        return None
                    
                    filepath = os.path.join(temp_dir, filename)
                    doc.save(filepath)
                    return (filename, filepath)
                except Exception as e:
                    return None
            
            # Запускаем параллельную генерацию
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(generate_row_file, row) for row in rows_data]
                
                for future in as_completed(futures):
                    result = future.result()
                    if result:
                        filename, filepath = result
                        zip_file.write(filepath, filename)
        
        zip_buffer.seek(0)
        response = FileResponse(
            zip_buffer, 
            as_attachment=True, 
            filename=f'Документы_ЕГРН_{current_date.strftime("%Y%m%d_%H%M%S")}.zip'
        )
        return response