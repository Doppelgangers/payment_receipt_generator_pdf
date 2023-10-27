import time
from io import BytesIO
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Mm

from barcode import Code39
from barcode.writer import ImageWriter
import qrcode


class PaymentReceipt:

    def __init__(self, path_template: str | Path):
        self.path_template = path_template

    @staticmethod
    def generate_qrcode(data):
        # Создаем изображение QR-кода
        qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=32, border=2)
        qr.add_data(data)
        qr.make(fit=True)
        qr_image = qr.make_image(fill_color="black", back_color="white")

        # Создаем объект BytesIO и сохраняем изображение QR-кода в него
        qr_bytes = BytesIO()
        qr_image.save(qr_bytes)
        qr_bytes.seek(0)
        return qr_bytes

    @staticmethod
    def generate_barcode(data):
        barcode_bytes = BytesIO()

        # Генерируем изображение штрихкода и сохраняем его
        Code39(data, writer=ImageWriter()).write(barcode_bytes)

        barcode_bytes.seek(0)
        return barcode_bytes

    def fill_docx_template(self, context, save_path_temp: str | Path = None) -> None:
        if save_path_temp is None: save_path_temp = "temp.docx"

        # Инициализация шаблона
        template = DocxTemplate(self.path_template)

        for data in context["items"]:
            qr_file = self.generate_qrcode(data["qr_data"])
            barcode_file = self.generate_barcode(data["barcode_data"])
            # data["barcode_image"] = InlineImage(template, barcode_file,  height=Mm(15), width=Mm(80))
            data["barcode_image"] = InlineImage(template, barcode_file,  height=Mm(15))
            data["qrcode_image"] = InlineImage(template, qr_file, width=Mm(28), height=Mm(28))
            data["total_sum"] = f"{int(data['total_sum'] // 1)} руб. {int(round(data['total_sum'] % 1, 2) * 100)} коп."


        template.render(context)
        # Сохранение pdf на основе шаблона
        template.save(save_path_temp)
        self.del_first_line_in_docx(save_path_temp)

    @staticmethod
    def del_first_line_in_docx(path_docx: str | Path):
        docx = Document(path_docx)
        # Удаляем первую строку

        docx._element.body.remove(docx.paragraphs[0]._element)
        # Сохраняем изменения
        docx.save(path_docx)


items = [
        {
            "qr_data": 'sdfgfjdpgjfpisjdgpifjdgoijfdopgsjdfopg23o3ijpoweifij23pofijojfpoijs',
            "barcode_data": '8546456459874568',
            "organization": 'МАДОУ "Детский сад № 100"',
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

        {
            "qr_data": 'qqweqweqweqweqweqwewqeqwewqeqweqweqwe',
            "barcode_data": '8746854859848484848484848456258',
            "organization": 'МАДОУ "Детский сад № 100"',
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

data = {"items": items.copy()}


if __name__ == '__main__':
    t = time.time()
    receipts = PaymentReceipt("templates/test.docx")
    receipts.fill_docx_template(data)
    print(time.time()-t)
