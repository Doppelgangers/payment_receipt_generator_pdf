import os
from pathlib import Path
from docxtpl import DocxTemplate
from docx2pdf import convert as docx_to_pdf


# @dataclass
# class ClientPaymentData:
#     full_name: str
#     client_personal_account: str  # Лицвеой счёт клиента
#     agreement_date: str  # Дата заключения договара
#
#
# @dataclass
# class OrganizationData:
#     organization: str
#     department: str
#     inn: str  # ИНН
#     kpp: str  # КПП
#     personal_account: str  # Лицвеой счёт учреждения
#     current_account: str  # Расчётный счёт учреждения
#     institution_address: str  # Адрес учреждения
#     bik: str  # БИК
#     correspondent_account: str  # Корреспондентский счёт
#
#
# @dataclass
# class PaymentData:
#     kbk: str  # Код бюджетной классификации
#     purpose_of_payment: str  # Назначение платежа
#     kind_of_activity: str  # Вид деятельности
#     date_payment: str  # За какой год и месяц произведена оплата
#     kindergarten_group: str  # Группа детского сада
#     total_sum: float
# test_data_org = OrganizationData(
#     organization="""МАДОУ "Детский сад № 100" """,
#     department="Департамент финансов г.Н.Новгорода",
#     inn="5260040678",
#     kpp="526001001",
#     personal_account="07040754581",
#     current_account="03234643227010003204",
#     institution_address="в ВОЛГО-ВЯТСКОЕ ГУ БАНКА РОССИИ//УФК по Нижегородской области г. Нижний Новгород",
#     bik="012202102",
#     correspondent_account="40102810745370000024",
# )
#
# test_data_cl = ClientPaymentData(
#     full_name="Тарасова Есения",
#     client_personal_account="4100100323",
#     agreement_date="01.10.20",
# )
#
# test_data_pl = PaymentData(
#     kbk="07507011130199404130",
#     purpose_of_payment="Оплата за Родительская плата за присмотр и уход за детьми.Май 2023 г./100 15 старшая 12ч",
#     date_payment="Май 2023 г",
#     kind_of_activity="04013",
#     total_sum=3193.20,
#     kindergarten_group="100 13 2 младшая"
# )


class PaymentReceipt:

    def __init__(self, path_template: str | Path):
        self.path_template = path_template

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
        docx_to_pdf("temp.docx", save_path)
        os.remove(temp_file)

data = {
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
        }

if __name__ == '__main__':
    receipt = PaymentReceipt("templates/void_template_1.docx")
    receipt.create_one(data, save_path="out.pdf", qrcode="bedcode.bmp", barcode="barcode.bmp")
