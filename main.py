import os
from pathlib import Path

from docx import Document
from docxtpl import DocxTemplate




class PaymentReceipt:

    def __init__(self, path_template: str | Path):
        self.path_template = path_template

    def create_multiply(self, context, save_path: str | Path = None) -> None:
        if save_path is None:
            save_path = "output.pdf"
        temp_file = "temp.docx"

        # Инициализация шаблона
        template = DocxTemplate(self.path_template)

        template.render(context)

        # Сохранение pdf на основе шаблона
        template.save(temp_file)
        self.del_first_line_in_docx(temp_file)

    @staticmethod
    def del_first_line_in_docx(path_docx: str | Path):
        docx = Document(path_docx)
        # Удаляем первую строку

        docx._element.body.remove(docx.paragraphs[0]._element)
        # Сохраняем изменения
        docx.save(path_docx)

    def create_one(self, context: dict, qrcode: str | Path, barcode: str | Path, save_path: str | Path = None) -> None:
        """
        :param barcode: Путь к файлу qr кода
        :param qrcode: Путь к файлу bar кода
        :param context:
            {
                 full_name: str Имя клиента
                 client_personal_account: str   Лицвеой счёт клиента
                 agreement_date: str  Дата заключения договара
                 organization: str Название организации
                 department: str Депортамент финансов
                 inn: str   ИНН
                 kpp: str   КПП
                 personal_account: str  Лицвеой счёт учреждения
                 current_account: str  Расчётный счёт учреждения
                 institution_address: str  Адрес учреждения
                 bik: str  БИК
                 correspondent_account: str  Корреспондентский счёт
                 kbk: str  Код бюджетной классификации
                 purpose_of_payment: str  Назначение платежа
                 kind_of_activity: str  Вид деятельности
                 date_payment: str  За какой год и месяц произведена оплата (Май 2023 г)
                 kindergarten_group: str  # Группа детского сада
                 total_sum: float Сумма платежа
            }

        :param save_path: путь куда сохранить pdf file
        """

        if save_path is None:
            save_path = "output.pdf"
        temp_file = "temp.docx"

        # Инициализация шаблона
        template = DocxTemplate(self.path_template)

        # Изменения в шаблоне
        context["total_sum"] = f"{int(context['total_sum']//1)} руб. {int(round(context['total_sum']%1,2)*100)} коп."
        template.replace_pic("qrcode_1", qrcode)
        template.replace_pic("barcode_1", barcode)
        template.render(context)

        #Сохранение pdf на основе шаблона
        template.save(temp_file)

        os.remove(temp_file)

items = [
        {
            "organization": """МАДОУ "Детский сад № 100" """,
            "department": "Департамент финансов г.Н.Новгорода",
            "inn": "5260040678",
            "kpp": "526001001",
            "personal_account": "07040754581",
            "current_account": "03234643227010003204",
            "institution_address": "в ВОЛГО-ВЯТСКОЕ ГУ БАНКА РОССИИ//УФК по Нижегородской области г. Нижний Новгород",
            "bik": "012202102",
            "correspondent_account": "40102810745370000024",
            "full_name": "Тарасова Есения",
            "client_personal_account": "4100100323",
            "agreement_date": "01.10.20",
            "kbk": "07507011130199404130",
            "purpose_of_payment": "Оплата за Родительская плата за присмотр и уход за детьми.",
            "date_payment": "Май 2023 г",
            "kind_of_activity": "04013",
            "total_sum": 3193.20,
            "kindergarten_group": "100 13 2 младшая"
        },
    ]

data = {"items": items*40}

if __name__ == '__main__':
    # receipt = PaymentReceipt("templates/void_template_1.docx")
    # receipt.create_one(items[0], save_path="out.pdf", qrcode="bedcode.bmp", barcode="barcode.bmp")

    receipts = PaymentReceipt("templates/multiply_template.docx")
    receipts.create_multiply(data)
