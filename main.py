import json
import os
import subprocess
import time
import datetime
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import qrcode
from barcode import Code39
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage


class PaymentReceiptData:
    class Money:
        def __init__(self, number_money: float):
            self.money = number_money

        @property
        def kopecks(self):
            return int(round(self.money % 1, 2) * 100)

        @property
        def rubles(self):
            return int(self.money)

        def __str__(self):
            return f"{self.rubles} руб. {self.kopecks} коп."

        @property
        def api_format(self):
            return f"{self.rubles}{self.kopecks}"

        def __repr__(self):
            return f"{self.__class__.__name__}({self.money})"

    class PDate:
        def __init__(self, str_date):
            self._date: datetime.date = self._serialize_date(str_date)

        @staticmethod
        def _serialize_date(str_date: str) -> datetime.date:
            """
            При изменении способа хранения даты в json стоит внести правки в этот метод класса
            :param str_date: строчный обхект даты
            :return: python.datetime.data object
            """
            day, month, year = map(lambda x: int(x), str_date.split("."))
            return datetime.date(year, month, day)

        @property
        def date_obj(self):
            return self._date

        @property
        def month_and_short_year(self) -> str:
            """
            Выводит дату в формете короткого года+месяц
            date(13.05.2023) -> 2305
            :return:
            """
            api_date = self._date.strftime("%Y%m")
            if len(api_date) == 6:
                api_date = api_date[2:]
            return api_date

        @property
        def month_word_and_year(self) -> str:
            """
            date(13.05.2023) -> Май 2023 г.
            """
            months = ['Январь', 'февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                      'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
            return f"{months[int(self.date_obj.month) - 1]} {int(self.date_obj.year)} г."

        @property
        def payment_period(self):
            """
            Выводит год месяц и вместо числа 00
            date(13.05.2023) -> 20230500
            """
            return self._date.strftime("%Y%m00")

        def __str__(self):
            return self._date.strftime("%d.%m.%Y")

        def __repr__(self):
            return f"{self.__class__.__name__}({self.__str__()})"

    def __init__(self, json_data):
        self.organization = None
        self.department = None
        self.inn = None
        self.kpp = None
        self.personal_account = None
        self.current_account = None
        self.bank_name = None
        self.bik = None
        self.correspondent_account = None
        self.full_name = None
        self.client_personal_account = None
        self.agreement_date: PaymentReceiptData.PDate = None
        self.kbk = None
        self.purpose_of_payment = None
        self.date_payment: PaymentReceiptData.PDate = None
        self.kind_of_activity = None
        self.total_sum: PaymentReceiptData.Money = None
        self.kindergarten_group = None

        # Инициализация атрибутов классада
        self.serialize_json(json_data)

    @property
    def name(self):
        return f"{self.department}({self.organization})"

    def serialize_json(self, json_str: str | dict):
        """
        Преобразует json в объекты python
        При изменении имён в json файле заменить их в json_dict.get(``` Новое имя ```)
        """
        if type(json_str) is str:
            json_dict: dict = json.loads(json_str)
        if type(json_str) is dict:
            json_dict: dict = json_str

        # Запорнение базовых атрибутов
        self.organization = json_dict.get("organization")
        self.department = json_dict.get("department")
        self.inn = json_dict.get("inn")
        self.kpp = json_dict.get("kpp")
        self.personal_account = json_dict.get("personal_account")
        self.current_account = json_dict.get("current_account")
        self.bank_name = json_dict.get("bank_name")
        self.bik = json_dict.get("bik")
        self.correspondent_account = json_dict.get("correspondent_account")
        self.full_name = json_dict.get("full_name")
        self.client_personal_account = json_dict.get("client_personal_account")
        self.agreement_date = json_dict.get("agreement_date")
        self.kbk = json_dict.get("kbk")
        self.purpose_of_payment = json_dict.get("purpose_of_payment")
        self.date_payment = json_dict.get("payment_period")
        self.kind_of_activity = json_dict.get("kind_of_activity")
        self.total_sum = json_dict.get("total_sum")
        self.kindergarten_group = json_dict.get("kindergarten_group")
        # Преобразование строчных атрибутов в объекты python
        self.total_sum = self.Money(self.total_sum)
        self.date_payment = self.PDate(self.date_payment)
        self.agreement_date = self.PDate(self.agreement_date)

    @property
    def context_item(self) -> dict:
        """
        Из этого словаря будет формироваться элемент шаблона,
        Здесь необходимо привести данные в тот str формат
        который ождается быть увиден в шаблоне
        :return:
        """
        context_item = {
            "organization": self.organization,
            "department": self.department,
            "inn": self.inn,
            "kpp": self.kpp,
            "personal_account": self.personal_account,
            "current_account": self.personal_account,
            "bank_name": self.bank_name,
            "bik": self.bik,
            "correspondent_account": self.correspondent_account,
            "full_name": self.full_name,
            "client_personal_account": self.client_personal_account,
            "agreement_date": self.agreement_date,
            "kbk": self.kbk,
            "purpose_of_payment": self.purpose_of_payment,
            "date_payment": self.date_payment.month_word_and_year,
            "kind_of_activity": self.kind_of_activity,
            "total_sum": self.total_sum.__str__(),
            "kindergarten_group": self.kindergarten_group
        }
        return context_item


class Codification:
    """
    Класс для создания Qrcode, Barcode и тд.
    """
    class QRCode:

        def __init__(self):
            self.qr_data: list[dict] = None

        def generate(self, qr_data: str = None) -> BytesIO:
            qr_data = qr_data if qr_data else self.qr_data
            # Создаем изображение QR-кода
            qr = qrcode.QRCode(version=1, error_correction=qrcode.constants.ERROR_CORRECT_L, box_size=32, border=2)
            qr.add_data(self.get_codification_str(qr_data))
            qr.make(fit=True)
            qr_image = qr.make_image(fill_color="black", back_color="white")
            # Создаем объект BytesIO и сохраняем изображение QR-кода в него
            qr_bytes = BytesIO()
            qr_image.save(qr_bytes)
            qr_bytes.seek(0)
            return qr_bytes

        def get_codification_str(self, attribute_list: [dict] = None):
            attribute_list = attribute_list if attribute_list else self.qr_data
            qr_data = 'ST00012|'
            for arg in attribute_list:
                qr_data += f'{arg["name"]}={arg["value"]}|'
            return qr_data

        def get_attribute_list_in_payment_data(self, payment_data: PaymentReceiptData) -> list[dict]:
            qr_data = \
                [
                    {"name": "Name", "value": payment_data.name},
                    {"name": "PersonalAcc", "value": payment_data.current_account},
                    {"name": "BankName", "value": payment_data.bank_name},
                    {"name": "BIC", "value": payment_data.bik},
                    {"name": "CorrespAcc", "value": payment_data.correspondent_account},
                    {"name": "PayeeINN", "value": payment_data.inn},
                    {"name": "PersonalAccount", "value": payment_data.personal_account},
                    {"name": "PersAcc", "value": payment_data.client_personal_account},
                    {"name": "Category", "value": payment_data.kind_of_activity},
                    {"name": "PaymPeriod", "value": payment_data.date_payment.payment_period},
                    {"name": "Sum", "value": payment_data.total_sum.api_format}
                ]
            self.qr_data = qr_data
            return qr_data

    class BarCode:

        def __init__(self, org_pres_acc, client_pers_acc, payment_period: PaymentReceiptData.PDate, summa: PaymentReceiptData.Money , kind_of_activity):
            self.org_pres_acc = org_pres_acc
            self.client_pers_acc = client_pers_acc
            self.payment_period = payment_period
            self.summa= summa
            self.kind_of_activity = kind_of_activity

        def generate(self) -> BytesIO:
            barcode_bytes = BytesIO()
            data_barcode = self.get_codification_str()
            # Генерируем изображение штрихкода и сохраняем его
            Code39(data_barcode, writer=ImageWriter()).write(barcode_bytes)
            barcode_bytes.seek(0)
            return barcode_bytes

        def get_codification_str(self) -> str:
            return f"00000{self.org_pres_acc}{self.client_pers_acc}{self.payment_period.month_and_short_year}0000{self.summa.api_format}{self.kind_of_activity}"

        def __str__(self):
            return self.get_codification_str()


class PaymentReceipt:

    def __init__(self, path_template: str | Path):
        self.path_template = path_template

    def render(self, list_json, save_path: str|Path = None, filename: str = None):
        filename = filename if filename else "receipt"
        save_path = Path('temp') if save_path is None else Path(save_path)
        self.fill_docx_template(list_json=list_json, save_path_file=save_path / (filename + ".docx"))
        try:
            self.convert_docx_to_pdf(file_path_docx=save_path/(filename+".docx"), save_folder_path_pdf=save_path, remove_docx=True)
        except Exception as e:
            os.remove(save_path/(filename+".docx"))
            raise Exception(e)

    def fill_docx_template(self, list_json: list[str], save_path_file: str | Path = None) -> None:
        if save_path_file is None: save_path_file = "temp.docx"
        # Инициализация шаблона
        template = DocxTemplate(self.path_template)
        items = []
        for json_data in list_json:
            data_payment = PaymentReceiptData(json_data)
            # Создание QR кода
            qr_code = Codification.QRCode()
            qr_code.get_attribute_list_in_payment_data(data_payment)
            qr_file = qr_code.generate()

            # Создание Bar кода
            bar_code = Codification.BarCode(
                org_pres_acc=data_payment.personal_account,
                client_pers_acc=data_payment.client_personal_account,
                payment_period=data_payment.date_payment,
                summa=data_payment.total_sum,
                kind_of_activity=data_payment.kind_of_activity,
            )
            barcode_file = bar_code.generate()

            context_item: dict = data_payment.context_item
            # Добавление элементов в контекст
            context_item["barcode_image"] = InlineImage(template, barcode_file, height=Mm(15))
            context_item["qrcode_image"] = InlineImage(template, qr_file, width=Mm(28), height=Mm(28))
            items.append(context_item)

        context = {"items": items}
        template.render(context)
        # Сохранение заполненного docx документа
        Path(save_path_file).parent.mkdir(parents=True, exist_ok=True)
        template.save(save_path_file)
        self.del_first_line_in_docx(save_path_file)

    @staticmethod
    def convert_docx_to_pdf(file_path_docx: str, save_folder_path_pdf: str, remove_docx=True):
        # Проверить, существует ли файл
        if not os.path.exists(file_path_docx):
            raise FileNotFoundError(f"Файл {file_path_docx} не найден.")
        # Получить имя файла без расширения
        file_name = os.path.splitext(file_path_docx)[0]
        # Конвертировать файл в pdf
        exit_code = os.system(f"soffice --headless --convert-to pdf {file_path_docx} --outdir {save_folder_path_pdf}")
        # Проверить код выхода команды
        if exit_code == 0:
            pdf_path = file_name + ".pdf"
            # Удалить docx-файл, если указано
            if remove_docx:
                os.remove(file_path_docx)
            return pdf_path
        else:
            raise Exception(f"Ошибка при конвертации файла {file_path_docx} в pdf. Возможно у вас не установлен libreoffice")

    @staticmethod
    def del_first_line_in_docx(path_docx: str | Path):
        """
        Удаляет лишнюю строчку после генерации файла по шаблону
        """
        docx = Document(path_docx)
        docx._element.body.remove(docx.paragraphs[0]._element)
        docx.save(path_docx)


if __name__ == '__main__':
    t = time.time()
    da = {
    "organization": "МАДОУ \"Детский сад № 100\"",
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
    "agreement_date": "01.10.2020",
    "kbk": "07507011130199404130",
    "purpose_of_payment": "Оплата за Родительская плата за присмотр и уход за детьми.",
    "payment_period": "01.05.2023",
    "kind_of_activity": "04013",
    "total_sum": 3193.20,
    "kindergarten_group": "100 13 2 младшая"
}

    receipts = PaymentReceipt("templates/template.docx")
    receipts.render([da]*5, save_path="D:\Desktop\sq", filename="test")
    print(time.time() - t)
