import time
import datetime
from io import BytesIO
from pathlib import Path

import qrcode
from barcode import Code128, Code39
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


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

    @staticmethod
    def _serialize_date(str_date: str) -> datetime:
        day, month, year = map(lambda x: int(x), str_date.split("."))
        return datetime.date(year, month, day)

    def fill_docx_template(self, context, save_path_temp: str | Path = None) -> None:
        if save_path_temp is None: save_path_temp = "temp.docx"

        # Инициализация шаблона
        template = DocxTemplate(self.path_template)

        for data in context["items"]:
            qr_data = self.generate_data_qrcode(name=f"{data['department']}({data['organization']})",
                current_account=data["current_account"],
                bank_name=data["bank_name"],
                bik=data["bik"],
                corresp_acc=data["correspondent_account"],
                inn=data["inn"],
                org_pres_acc=data["personal_account"],
                client_pers_acc=data["client_personal_account"],
                kind_of_activity=data["kind_of_activity"],
                summa=data["total_sum"]
            )
            qr_file = self.generate_qrcode(qr_data)

            barcode_data = self.generate_data_barcode(
                org_pres_acc=data["personal_account"],
                client_pers_acc=data["client_personal_account"],
                date_payment=self._serialize_date(data["date_payment"]),
                summa=data["total_sum"],
                kind_of_activity=data["kind_of_activity"]
            )
            barcode_file = self.generate_barcode(barcode_data)

            months = ['Январь', 'февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                      'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
            date_payment = self._serialize_date(data["date_payment"])

            data["date_payment"] = f"{months[int(date_payment.month)-1]} {int(date_payment.year)} г."
            data["barcode_image"] = InlineImage(template, barcode_file,  height=Mm(15))
            data["qrcode_image"] = InlineImage(template, qr_file, width=Mm(28), height=Mm(28))
            data["total_sum"] = f"{int(data['total_sum'])} руб. {int(round(data['total_sum'] % 1, 2) * 100)} коп."

        template.render(context)
        # Сохранение pdf на основе шаблона
        template.save(save_path_temp)
        self.del_first_line_in_docx(save_path_temp)

    def generate_data_barcode(self, org_pres_acc: str, client_pers_acc: str, date_payment: datetime.date, summa: float, kind_of_activity: str):
        month = str(date_payment.month)
        if len(month) == 1: month = "0" + month
        year = str(date_payment.year)
        year = year[2:] if (len(year) ==4) else year
        summa = f"{int(summa)}{int(round(summa % 1, 2) * 100)}"
        template = f"00000{org_pres_acc}{client_pers_acc}{month}{year}0000{summa}{kind_of_activity}"
        return template

    @staticmethod
    def generate_data_qrcode(**kwargs):
        name: str = kwargs.get("name")
        current_account: str = kwargs.get("current_account")
        bank_name: str = kwargs.get("bank_name")
        bik: str = kwargs.get("bik")
        corresp_acc: str = kwargs.get("corresp_acc")
        inn: str = kwargs.get("inn")
        org_pres_acc: str = kwargs.get("org_pres_acc")
        client_pers_acc: str = kwargs.get("client_pers_acc")
        kind_of_activity: str = kwargs.get("kind_of_activity")
        summa: float = kwargs.get("summa")
        summa = f"{int(summa)}{int(round(summa % 1, 2) * 100)}"
        template = f"ST00012|Name={name}|PersonalAcc={current_account}|BankName={bank_name}|BIC={bik}|CorrespAcc={corresp_acc}|PayeeINN={inn}|PersonalAccount={org_pres_acc}|PersAcc={client_pers_acc}|Category={kind_of_activity}|PaymPeriod=20230500|Sum={summa}|"
        return template

    @staticmethod
    def del_first_line_in_docx(path_docx: str | Path):
        docx = Document(path_docx)
        # Удаляем первую строку
        docx._element.body.remove(docx.paragraphs[0]._element)
        # Сохраняем изменения
        docx.save(path_docx)


items = [
        {
            "organization": 'МАДОУ "Детский сад № 100"',
            "department": "Департамент финансов г.Н.Новгорода",
            "inn": "5260040678",
            "kpp": "526001001",
            "personal_account": "07040754581",
            "current_account": "03234643227010003204",
            "bank_name": "ВОЛГО-ВЯТСКОЕ ГУ БАНКА РОССИИ//УФК по Нижегородской области г. Нижний Новгород",
            "bik": "012202102",
            "correspondent_account": "40102810745370000024",
            "full_name": "Тарасова Есения",
            "client_personal_account": "4100100232",
            "agreement_date": "01.10.20",
            "kbk": "07507011130199404130",
            "purpose_of_payment": "Оплата за Родительская плата за присмотр и уход за детьми.",
            "date_payment": "01.05.2023",
            "kind_of_activity": "04013",
            "total_sum": 3193.20,
            "kindergarten_group": "100 13 2 младшая"
        },

        {
            "organization": 'МАДОУ "Детский сад № 100"',
            "department": "Департамент финансов г.Н.Новгорода",
            "inn": "5260040678",
            "kpp": "526001001",
            "personal_account": "07040754581",
            "current_account": "03234643227010003204",
            "bank_name": "ВОЛГО-ВЯТСКОЕ ГУ БАНКА РОССИИ//УФК по Нижегородской области г. Нижний Новгород",
            "bik": "012202102",
            "correspondent_account": "40102810745370000024",
            "full_name": "Тарасова Есения",
            "client_personal_account": "4100100323",
            "agreement_date": "01.10.20",
            "kbk": "07507011130199404130",
            "purpose_of_payment": "Оплата за Родительская плата за присмотр и уход за детьми.",
            "date_payment": "01.05.23",
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
