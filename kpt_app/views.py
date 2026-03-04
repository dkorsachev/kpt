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
    
    # Получаем данные о строках из формы
    rows_data = form.cleaned_data['rows_json']
    current_date = datetime.now()
    
    # Пути к шаблонам
    templates_dir = os.path.join(settings.BASE_DIR, 'templates_docs')
    
    # Проверяем различные варианты имен файлов
    sluzhebka_template = os.path.join(templates_dir, 'sluzhebka_template.docx')
    if not os.path.exists(sluzhebka_template):
        sluzhebka_template = os.path.join(templates_dir, 'sluzhebka_template.docx.docx')
    
    kpt_template = os.path.join(templates_dir, 'kpt_template.docx')
    if not os.path.exists(kpt_template):
        kpt_template = os.path.join(templates_dir, 'kpt_template.docx.docx')
    
    zu_template = os.path.join(templates_dir, 'zu_template.docx')
    if not os.path.exists(zu_template):
        zu_template = os.path.join(templates_dir, 'zu_template.docx.docx')
    
    # Проверяем существование файлов шаблонов
    if not os.path.exists(sluzhebka_template):
        return render(request, 'kpt_app/index.html', {
            'form': form,
            'error': f'Файл шаблона не найден: sluzhebka_template.docx'
        })
    
    # Создаем генератор документов
    generator = DocumentGenerator()
    
    # Создаем временную директорию для файлов
    with tempfile.TemporaryDirectory() as temp_dir:
        # Создаем ZIP-архив в памяти
        zip_buffer = BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Генерируем служебную записку (DOCX и PDF)
            try:
                sluzhebka_doc = generator.generate_sluzhebka(
                    sluzhebka_template,
                    rows_data,
                    current_date
                )
                sluzhebka_docx = os.path.join(temp_dir, "Запрос из ЕГРН.docx")
                sluzhebka_doc.save(sluzhebka_docx)
                
                # Добавляем в архив DOCX
                zip_file.write(sluzhebka_docx, "Запрос из ЕГРН.docx")
                
                # Конвертируем в PDF и добавляем в архив
                sluzhebka_pdf = os.path.join(temp_dir, "Запрос из ЕГРН.pdf")
                if generator.convert_to_pdf(sluzhebka_docx, sluzhebka_pdf):
                    zip_file.write(sluzhebka_pdf, "Запрос из ЕГРН.pdf")
            except Exception as e:
                return render(request, 'kpt_app/index.html', {
                    'form': form,
                    'error': f'Ошибка при генерации служебной записки: {str(e)}'
                })
            
            # Генерируем отдельные файлы для каждой строки (только DOCX)
            for i, row in enumerate(rows_data):
                try:
                    cad_num = row['cadastral_number']
                    contract_number = row['contract_number']
                    contract_date = datetime.strptime(row['contract_date'], '%Y-%m-%d')
                    
                    # Определяем тип кадастрового номера
                    if re.match(r'^\d{2}:\d{2}:\d{7}$', cad_num):
                        # Кадастровый квартал
                        if not os.path.exists(kpt_template):
                            continue
                        doc = generator.generate_kpt(
                            kpt_template,
                            cad_num,
                            contract_number,
                            contract_date,
                            current_date
                        )
                        # Формируем имя файла
                        num_parts = cad_num.replace(':', '_')
                        filename = f"Запрос_на_кад_квартал_{num_parts}.docx"
                    else:
                        # Земельный участок
                        if not os.path.exists(zu_template):
                            continue
                        doc = generator.generate_zu(
                            zu_template,
                            cad_num,
                            contract_number,
                            contract_date,
                            current_date
                        )
                        # Формируем имя файла
                        num_parts = cad_num.replace(':', '_')
                        filename = f"Запрос_выписки_ЗУ_{num_parts}.docx"
                    
                    # Сохраняем DOCX
                    docx_path = os.path.join(temp_dir, filename)
                    doc.save(docx_path)
                    
                    # Добавляем только DOCX в архив (без PDF)
                    zip_file.write(docx_path, filename)
                        
                except Exception as e:
                    # Если ошибка с конкретным номером, пропускаем его
                    continue
        
        # Возвращаем ZIP-архив
        zip_buffer.seek(0)
        response = FileResponse(
            zip_buffer, 
            as_attachment=True, 
            filename=f'Документы_ЕГРН_{current_date.strftime("%Y%m%d_%H%M%S")}.zip'
        )
        return response