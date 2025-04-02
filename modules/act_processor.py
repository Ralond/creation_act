import os
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
from modules.file_manager import FileManager
from config import DEFAULT_TEMPLATE, DEFAULT_SOURCE, DEFAULT_REGISTER

class ActProcessor:
    def __init__(self):
        self.file_manager = FileManager()

    def validate_row(self, row):
        """Проверка валидности строки реестра"""
        if not isinstance(row, (list, tuple)):
            return False
        
        if all(cell is None for cell in row):
            return False
        
        if len(row) < 7:
            return False
        
        if not (row[1] and row[6] and isinstance(row[6], datetime)):
            return False
        
        return True

    def process_register(self, register_path):
        """Обработка реестра актов"""
        wb = self.file_manager.load_workbook_safe(register_path)
        sheet = wb['Реестр']
        
        valid_rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if self.validate_row(row):
                valid_rows.append(row)
            elif any(cell is not None for cell in row):
                break  # Прекращаем при первой частично заполненной строке
                
        wb.close()
        return valid_rows

    def generate_akt(self, row, template_path, source_path, output_dir):
        """Генерация одного акта с полной логикой"""
        akt_num = f"{self.file_manager.safe_get(row, 0, '')}{self.file_manager.safe_get(row, 1, '')}"
        akt_date = self.file_manager.safe_get(row, 6)
        
        # Создаем папку и файл для акта
        akt_folder = self.file_manager.create_akt_folder(akt_num)
        akt_file = akt_folder / f"Акт_{akt_num}.xlsx"
        
        try:
            if not akt_date:
                raise ValueError("Дата акта не указана")

            # Загружаем необходимые данные
            wb_source = self.file_manager.load_workbook_safe(source_path)
            wb_template = self.file_manager.load_workbook_safe(template_path)
            
            source_sheet = wb_source['Реквизиты']
            involved_sheet = wb_source['Причастные']
            template_sheet = wb_template['1']
            
            # Создаем новую книгу для акта
            output_wb = openpyxl.Workbook()
            output_wb.remove(output_wb.active)
            
            # Создаем лист для акта
            sheet_name = f"Акт {akt_num}"[:31]
            new_sheet = output_wb.create_sheet(title=sheet_name)
            
            # Копируем шаблон
            self._copy_template(template_sheet, new_sheet)
            
            # Получаем данные о лицах
            persons = self._get_persons_data(involved_sheet, akt_date)
            
            # Заполняем основные данные
            self._fill_main_data(new_sheet, row, persons)
            
            # Заполняем информацию о проекте
            project_info = self._format_project_info(row)
            if project_info:
                self._write_to_cell(new_sheet, 'R52', project_info)
            
            # Сохраняем файл акта
            self.file_manager.save_workbook_safe(output_wb, akt_file)
            
            return {
                'folder': akt_folder,
                'file': akt_file,
                'akt_num': akt_num,
                'akt_date': akt_date.strftime('%d.%m.%Y'),
                'status': 'success'
            }
            
        except Exception as e:
            return {
                'folder': akt_folder,
                'file': akt_file,
                'akt_num': akt_num,
                'error': str(e),
                'status': 'error'
            }
        finally:
            # Закрываем workbook если они были открыты
            if 'wb_source' in locals():
                wb_source.close()
            if 'wb_template' in locals():
                wb_template.close()
            if 'output_wb' in locals():
                output_wb.close()

    def _copy_template(self, template_sheet, new_sheet):
        """Копирование шаблона с сохранением стилей"""
        # Копируем объединенные ячейки
        for merged_range in template_sheet.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_range))
        
        # Копируем размеры столбцов
        for col in template_sheet.columns:
            col_letter = get_column_letter(col[0].column)
            if col_letter in template_sheet.column_dimensions:
                new_sheet.column_dimensions[col_letter].width = template_sheet.column_dimensions[col_letter].width
        
        # Копируем высоты строк
        for row_idx, row_dim in template_sheet.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = row_dim.height
        
        # Копируем ячейки с сохранением стилей
        for row in template_sheet.iter_rows():
            for cell in row:
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                    
                new_cell = new_sheet[cell.coordinate]
                new_cell.value = cell.value
                
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

    def _get_persons_data(self, involved_sheet, akt_date):
        """Получение данных о причастных лицах"""
        return {
            'client_tech_supervision': self._get_most_recent_person(involved_sheet, akt_date, 4, 13),
            'general_contractor': self._get_most_recent_person(involved_sheet, akt_date, 15, 24),
            'contractor_tech_supervision': self._get_most_recent_person(involved_sheet, akt_date, 26, 35),
            'author_supervision': self._get_most_recent_person(involved_sheet, akt_date, 37, 46),
            'work_executor': self._get_most_recent_person(involved_sheet, akt_date, 59, 68),
            'others': self._get_most_recent_person(involved_sheet, akt_date, 70, 79)
        }

    def _get_most_recent_person(self, sheet, akt_date, start_row, end_row):
        """Поиск актуальных данных на дату акта"""
        most_recent_data = {}
        most_recent_date = None
        
        for row_num in range(start_row, end_row + 1):
            try:
                order_date = sheet.cell(row=row_num, column=1).value
                if not order_date or not isinstance(order_date, datetime):
                    continue
                    
                if order_date <= akt_date:
                    if most_recent_date is None or order_date > most_recent_date:
                        most_recent_data = {
                            'name': sheet.cell(row=row_num, column=6).value,
                            'position': sheet.cell(row=row_num, column=2).value,
                            'organization': sheet.cell(row=row_num, column=3).value,
                            'order': sheet.cell(row=row_num, column=4).value,
                            'nrs': sheet.cell(row=row_num, column=5).value,
                            'address': sheet.cell(row=row_num, column=7).value
                        }
                        most_recent_date = order_date
            except Exception:
                continue
        
        return most_recent_data

    def _fill_main_data(self, sheet, row, persons):
        """Заполнение основных данных акта"""
        # Основные данные
        mapping = {
            'B26': f"{self.file_manager.safe_get(row, 0, '')}{self.file_manager.safe_get(row, 1, '')}",
            'AB26': self.file_manager.safe_get(row, 6),
            'A50': self.file_manager.safe_get(row, 2, ''),
            'M61': self.file_manager.safe_get(row, 4, ''),
            'M62': self.file_manager.safe_get(row, 5, ''),
            'A67': self.file_manager.safe_get(row, 7, ''),
            'A56': self.file_manager.safe_get(row, 8, ''),
            'A59': self.file_manager.safe_get(row, 9, ''),
            'G70': self.file_manager.safe_get(row, 13, ''),
            'J69': self.file_manager.safe_get(row, 14, ''),
            'A72': self._format_attachments(row)
        }
        
        for coord, value in mapping.items():
            self._write_to_cell(sheet, coord, value)
        
        # Данные о лицах (верхняя часть)
        upper_mapping = {
            'A28': self._format_person_details(persons.get('client_tech_supervision', {}), True),
            'A31': self._format_person_details(persons.get('general_contractor', {})),
            'A34': self._format_person_details(persons.get('contractor_tech_supervision', {})),
            'A37': self._format_person_details(persons.get('author_supervision', {})),
            'A40': self._format_person_details(persons.get('work_executor', {})),
            'A43': self._format_person_details(persons.get('others', {}))
        }
        
        for coord, value in upper_mapping.items():
            if value:
                self._write_to_cell(sheet, coord, value)
        
        # ФИО (нижняя часть)
        lower_mapping = {
            'A75': persons.get('client_tech_supervision', {}).get('name'),
            'A78': persons.get('general_contractor', {}).get('name'),
            'A81': persons.get('contractor_tech_supervision', {}).get('name'),
            'A84': persons.get('author_supervision', {}).get('name'),
            'A87': persons.get('work_executor', {}).get('name'),
            'A90': persons.get('others', {}).get('name')
        }
        
        for coord, value in lower_mapping.items():
            if value:
                self._write_to_cell(sheet, coord, value)
        
        # Регулировка высоты строк
        self._adjust_row_heights(sheet)

    def _write_to_cell(self, sheet, coord, value):
        """Безопасная запись в ячейку"""
        cell = sheet[coord]
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            cell.value = value
            return True
        return False

    def _adjust_row_heights(self, worksheet):
        """Автоподбор высоты строк"""
        for row in worksheet.iter_rows():
            max_lines = 1
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    max_lines = max(max_lines, cell.value.count('\n') + 1)
            
            if max_lines > 1:
                worksheet.row_dimensions[row[0].row].height = 15 * max_lines

    def _format_project_info(self, row):
        """Форматирование информации о проекте"""
        project_num = self.file_manager.safe_get(row, 10, '')
        project_sheet = self.file_manager.safe_get(row, 11, '')
        
        parts = []
        if project_num: parts.append(f"№ {project_num}")
        if project_sheet: parts.append(f"лист {project_sheet}")
        
        return "\n".join(parts) if parts else None

    def _format_attachments(self, row):
        """Форматирование приложений"""
        schemes = self.file_manager.safe_get(row, 9, '')
        materials = self.file_manager.safe_get(row, 8, '')
        
        parts = []
        if schemes: parts.append(str(schemes))
        if materials and materials != 'Не использовались':
            parts.append(f"Сертификаты к материалам: {materials}")
        
        return "\n".join(parts) if parts else None

    def _format_person_details(self, data, include_address=False):
        """Форматирование данных о лице"""
        if not data or not data.get('name'):
            return None
            
        parts = []
        if data.get('position'): parts.append(data['position'])
        if data.get('organization'): parts.append(data['organization'])
        if include_address and data.get('address'): parts.append(data['address'])
        
        details = []
        if data.get('order'): details.append(f"Приказ: {data['order']}")
        if data.get('nrs'): details.append(f"НРС: {data['nrs']}")
        
        result = ', '.join(filter(None, parts))
        if details:
            result += f" ({'; '.join(details)})"
        
        return result if result.strip() else None