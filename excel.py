#!/usr/bin/python
import tempfile
from asyncio.tasks import shield
from datetime import datetime
from locale import setlocale, LC_ALL
from os.path import dirname, join, exists, basename, splitext
from os import mkdir

from subprocess import run
import asyncio

from docx.shared import Mm
from pandas import read_excel as pd, to_datetime
from openpyxl import load_workbook
from docxtpl import DocxTemplate, InlineImage
from pdf2image import convert_from_path
from pymorphy2 import MorphAnalyzer
from openpyxl_image_loader import SheetImageLoader


def convert_to_pdf(file: str):
    run(f'unoconv -f pdf {file}')


class File:

    def __init__(self):
        self.path = dirname(__file__)
        self.locale = setlocale(LC_ALL, 'ru_RU')
        self.morph = MorphAnalyzer()
        self.read_file = pd(join(self.path, 'Регистрация.xlsx'), na_filter=False).to_dict('index')
        self.count = len(self.read_file) - 1

    async def make_folder(self, folder: str):
        print(f'Создается папка {folder}')
        if not exists(join(self.path, f'{folder}')):
            mkdir(join(self.path, f'{folder}'))
        return join(self.path, f'{folder}')

    async def get_image(self):
        print(f'Получение образца подписи')
        for img in range(self.count):
            image_loader = SheetImageLoader(load_workbook(join(self.path, 'Регистрация.xlsx')).worksheets[0])
            image = image_loader.get(f'B{2 + img}')
            image.save(join(self.path, 'files', 'TEMP', f'B{2 + img}.png'))

    async def get_context(self, i_dict: dict) -> dict:
        context = {
            "НАЗВАНИЕ": i_dict.get("Фирма").strip(),
            "ЮР_ГОР": i_dict.get("Юр. Адрес"),
            "ГОРОД": i_dict.get("Юр.Город"),
            "ДАТА": datetime.today().strftime("«%d» %B %Y года."),
            "ФИО": i_dict.get("ФИО"),
            "Ф_СОКР": f"{i_dict.get('ФИО').split(' ')[0]} {i_dict.get('ФИО').split(' ')[1][:1]}.{i_dict.get('ФИО').split(' ')[2][:1]}.",
            "НОМЕР_УСТАВА": i_dict.get("НОМЕР УСТАВА"),
            "ИНН": i_dict.get("ИНН"),
            "СУММ_ПРО": i_dict.get("Уставной Капитал"),
        }
        if i_dict["ИНН2"]:
            context["ДАТА_ПОД"] = to_datetime(i_dict.get("ДАТА ПОДАЧИ")).strftime("«%d» %B %Y г.")
            context["ДАТА"] = to_datetime(i_dict.get("ДАТА ПОДАЧИ")).strftime("«%d» %B %Y г.")
            context["ДАТА_РЕГ"] = to_datetime(i_dict.get("Дата Регистрации")).strftime("«%d» %B %Y г.")
        return context

    async def make_resheie(self, context: dict, image: int, file='Reshenie.docx'):
        file_path = await self.make_folder(context.get('НАЗВАНИЕ'))
        write_resh = DocxTemplate(join(self.path, 'files', 'templates', file))
        context["Image"] = InlineImage(write_resh, join(self.path, 'files', 'TEMP', f"B{2 + int(image)}.png"),
                                       width=Mm(40))
        write_resh.render(context=context)
        if file == "Reshenie.docx":

            file = f'РЕШЕНИЕ_{context["НАЗВАНИЕ"]}.docx'
        else:
            file = f'ПРИКАЗ_{context["НАЗВАНИЕ"]}.docx'
        write_resh.save(join(file_path, file))
        return join(file_path, file)

    async def start_cmd(self, cmd):
        proc = await asyncio.create_subprocess_shell(
            cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE)

        stdout, stderr = await proc.communicate()
        if stderr:
            print(stderr.decode())

    async def convert_to_pdf(self, file: str):
        file_pdf = file.split('_')[1].split('.')[0]
        await shield(self.start_cmd(f'unoconv -f pdf {file}'))
        return file_pdf

    async def convert_to_tiff(self, file):

        file_pdf_dir = join(self.path, file, f'РЕШЕНИЕ_{file}.pdf')
        with tempfile.TemporaryDirectory() as path:
            images_from_path = convert_from_path(
                file_pdf_dir,
                output_folder=path,
                last_page=1,
                dpi=300,
                first_page=0,
                grayscale=True
            )
        base_filename = splitext(basename(file_pdf_dir))[0] + '.tiff'
        tiff_dir = join(self.path, file, base_filename)
        for page in images_from_path:
            page.save(tiff_dir, 'JPEG')

    async def run(self):
        await self.get_image()
        for i_dict in range(self.count):
            context = await self.get_context(self.read_file[i_dict])
            docx = await self.make_resheie(context=context, image=i_dict)
            await self.convert_to_tiff(await self.convert_to_pdf(docx))
            if self.read_file[i_dict]['ИНН2']:
                docx2 = await self.make_resheie(context=context, image=i_dict, file='Prikaz.docx')
                await self.convert_to_pdf(docx2)


if __name__ == '__main__':
    asyncio.run(File().run())
