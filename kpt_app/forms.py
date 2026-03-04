from django import forms
import re
import json


class CadastralNumberForm(forms.Form):
    # Поле для JSON с данными о строках (кадастровый номер + договор)
    rows_json = forms.CharField(
        widget=forms.HiddenInput(),
        required=False,
        initial='[]'
    )
    
    def clean_rows_json(self):
        data = self.cleaned_data.get('rows_json', '[]')
        try:
            rows = json.loads(data)
            if not rows:
                raise forms.ValidationError('Добавьте хотя бы одну строку')
            
            # Проверка формата кадастровых номеров
            pattern_quarter = r'^\d{2}:\d{2}:\d{7}$'  # 95:06:1501006
            pattern_land = r'^\d{2}:\d{2}:\d{7}:\d{1,5}$'  # 95:09:0100041:223
            
            for row in rows:
                cad_num = row.get('cadastral_number', '')
                if not (re.match(pattern_quarter, cad_num) or re.match(pattern_land, cad_num)):
                    raise forms.ValidationError(
                        f'Неверный формат кадастрового номера: {cad_num}. '
                        'Ожидается формат XX:XX:XXXXXXX или XX:XX:XXXXXXX:XXXXX'
                    )
                
                if not row.get('contract_number'):
                    raise forms.ValidationError('Номер договора обязателен')
                    
                if not row.get('contract_date'):
                    raise forms.ValidationError('Дата договора обязательна')
            
            return rows
        except json.JSONDecodeError:
            raise forms.ValidationError('Ошибка в данных формы')