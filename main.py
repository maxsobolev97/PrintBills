import win32com.client as com
import configparser
import os
import zipfile
import subprocess
import PyPDF2
import pikepdf
import datetime as date
import shutil
import time
from docxtpl import DocxTemplate


class conf:

    def __init__(self):
        self.file = 'test.ini'

    def ReadPathConfig(self):
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        billPath = config['DEFAULT']['path']
        return billPath

    def ReadRarConfig(self):
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        rarPath = config['DEFAULT']['rar']
        return rarPath

    def ReadSectionConfig(self):
        sections = []
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        for section in config.sections():
            sections.append(config[section])
        return sections

    def ReadEmailsConfig(self):
        sections = []
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        for section in config.sections():
            sections.append([section, config[section]['email']])
        return sections

    def ReadPageConfig(self, section):
        billpages = []
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        settings = config[section]
        for setting in settings:
            if setting.__contains__('filesettings'):
                filesettings = config[section][setting]
                billpages.append(filesettings.split(','))
        return billpages

    def ReadAllConfig(self):
        config = configparser.ConfigParser()
        config.read(self.file, 'utf_8')
        return config


class mail:
    def __init__(self, config):
        self.Outlook = com.Dispatch('outlook.application')
        self.Namespace = self.Outlook.GetNameSpace("MAPI")
        self.InboxFolder = self.Namespace.GetDefaultFolder(6)
        self.InboxItems = self.InboxFolder.items
        self.config = config
        maildirs = self.config.ReadSectionConfig()
        for maildir in maildirs:
            try:
                self.InboxFolder.Folders.Item('Счета').Folders.Add(maildir.name)
            except:
                pass

    def findBills(self):
        emails = self.config.ReadEmailsConfig()
        print(f'Начинаю проверку папки входящие...')
        for InboxItem in self.InboxItems:
            ItemSubject = InboxItem.Subject
            ItemSender = InboxItem.SenderEmailAddress
            print(f'Проверка письма "{ItemSubject}"')
            for email in emails:
                if ItemSender in email:
                    print(f'Обнаружено письмо от {email[0]}')
                    path = self.config.ReadPathConfig()
                    billPath = email[0]
                    print(f'Сохранение вложения из письма "{ItemSubject}"')
                    self.downloadBill(InboxItem, billPath, path)
                    print(f'Перенос письма в папку {email[0]}')
                    self.moveInbox(InboxItem, billPath)

    def downloadBill(self, InboxItem, dirPath, mainPath):
        attaches = InboxItem.Attachments
        if len(attaches) > 0:
            for attach in attaches:
                timestamp = date.datetime.now().microsecond
                path = mainPath + '/' + dirPath + '/' + str(timestamp) + '_' + attach.FileName
                attach.SaveAsFile(path)

    def moveInbox(self, InboxItem, targetFolder):
        folder = self.InboxFolder.Folders.Item("Счета").Folders.Item(targetFolder)
        InboxItem.Move(folder)


class files:
    def __init__(self, config):
        self.path = config.ReadPathConfig()
        self.rar = config.ReadRarConfig()
        dirs = config.ReadSectionConfig()
        for dir in dirs:
            dirpath = str(self.path) + '/' + dir.name
            arcpath = dirpath + '/Архив'
            try:
                os.mkdir(dirpath)
                os.mkdir(arcpath)
            except:
                pass

    def billdirs(self):
        return os.listdir(self.path)

    def billfiles(self, dirname):
        path = self.path + '/' + dirname
        return os.listdir(path)

    def billpath(self, file, dir):
        if file != 'Архив' and file != '!Шаблон для согласования.doc':
            filepath = self.path + '/' + dir + '/' + file
            return filepath

    def expandArchive(self, filepath):
        print(f'Обнаружен нераспакованный архив')
        extracted = os.path.dirname(filepath)
        if filepath.lower().__contains__('.zip'):
            print(f'Распаковка архива')
            with zipfile.ZipFile(filepath, 'r') as zip:
                zip.extractall(extracted)
            self.movetoarc(filepath)
            files = os.listdir(extracted)
            return files.remove('Архив')
        elif filepath.lower().__contains__('.rar'):
            print(f'Распаковка архива')
            args = [str(self.rar), 'e', str(filepath), str(extracted)]
            subprocess.run(args)
            os.remove(filepath)
            files = os.listdir(extracted)
            return files.remove('Архив')

    def extractallarchives(self):
        print(f'Начинаю проверку папок на нераспакованные архивы...')
        dirs = self.billdirs()
        for dir in dirs:
            bills = self.billfiles(dir)
            for bill in bills:
                billpath = self.billpath(bill, dir)
                if billpath:
                    if billpath.lower().__contains__('.zip') or billpath.lower().__contains__('.rar'):
                        extracted = files.expandArchive(billpath)

    def movetoarc(self, filepath):
        timestamp = date.datetime.now().microsecond
        dstpath = os.path.dirname(filepath)
        dstname = os.path.basename(filepath)
        print(f'Перенос файла {dstname} в архив')
        dstfile = dstpath + '/Архив/' + str(timestamp) + '_' + dstname
        shutil.move(filepath, dstfile)


class documents:
    def __init__(self, config, files):
        self.config = config
        self.files = files
        self.templatePath = 'template.docx'

    def WordToPdf(self, filepath):
        encodedFilepath = filepath.replace('/', '\\')
        basename = os.path.basename(filepath)
        extension = basename.split('.')
        dirname = os.path.dirname(filepath)
        pdfName = dirname + '/' + extension[0] + '.pdf'
        pdfNewName = pdfName.replace('/', '\\')
        try:
            print('Попытка конвертации WORD в PDF')
            word = com.Dispatch('Word.Application')
            document = word.Documents.Open(encodedFilepath)
            document.SaveAs(pdfNewName, 17)
            document.Close()
            word.Quit()
            print('Документ успешно сконвертирован в PDF')
            self.files.movetoarc(filepath)
            return pdfName
        except:
            print('Не сработала конвертация из WORD в PDF!')
            word.Quit()

    def WordDocTypes(self):
        types = ['rtf', 'doc', 'docx']
        return types

    def makeagreement(self, filepath):
        print('Выполняется генерация листа согласования')
        dirname = os.path.dirname(filepath)
        basename = os.path.basename(filepath)
        section = os.path.basename(dirname)
        allsettings = self.config.ReadAllConfig()
        filesettings = allsettings[section]
        for setting in filesettings:
            if setting.__contains__('filesettings'):
                if filesettings[setting].split(',')[0] in basename:
                    newfilename = filepath + '.docx'
                    textdocid = setting[-1]
                    textlabel = 'approval' + str(textdocid)
                    text = filesettings[textlabel]
                    doc = DocxTemplate(self.templatePath)
                    text = {'text': text}
                    doc.render(context=text)
                    doc.save(newfilename)
                    time.sleep(3)
                    self.printdocument(newfilename)

    def makefiletoprint(self, filepath):
        dirname = os.path.dirname(filepath)
        section = os.path.basename(dirname)
        filesettings = self.config.ReadPageConfig(section)
        pdfFileObj = open(filepath, 'rb')
        try:
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        except Exception as e:
            pdfFileObj.close()
            pdfFileToRepair = pikepdf.Pdf.open(filepath)
            RepairedPath = os.path.dirname(filepath)
            RepairedNames = os.path.basename(filepath).split('.')
            RepairedName = RepairedPath + '/' + RepairedNames[0] + '_repaired.' + RepairedNames[1]
            pdfFileToRepair.save(RepairedName)
            pdfFileToRepair.close()
            time.sleep(3)
            os.remove(filepath)
            try:
                pdfReader = PyPDF2.PdfFileReader(RepairedName)
            except Exception as e:
                print(e)
                pdfFileObj.close()
        pdfWriter = PyPDF2.PdfFileWriter()
        newFilePath = os.path.dirname(filepath)
        filenames = os.path.basename(filepath).split('.')
        print(f'Начинаю конвертацию файла {filenames[0]} по заданному шаблону...')
        newFileName = newFilePath + '/' + filenames[0] + '_new.' + filenames[1]
        rangetoprint = []
        if len(filesettings) == 1:
            for setting in filesettings:
                rangetoprint = setting[1:]
        elif len(filesettings) == 0:
            rangetoprint = pdfReader.numPages
        elif len(filesettings) > 1:
            for setting in filesettings:
                if setting[0] in filenames[0]:
                    rangetoprint = setting[1:]

        if rangetoprint:
            for pageNum in rangetoprint:
                print(f'Вытаскиваю страницы...')
                pageObj = pdfReader.getPage(int(pageNum) - 1)
                pdfWriter.addPage(pageObj)
            with open(newFileName, "wb") as output:
                print(f'Сохранияю новый файл')
                pdfWriter.write(output)
            pdfFileObj.close()
            self.files.movetoarc(filepath)
            return newFileName
        else:
            pdfFileObj.close()
            self.files.movetoarc(filepath)

    def printdocument(self, filepath):
        print(f'Печать документа')
        os.startfile(filepath, 'print')
        time.sleep(3)
        self.files.movetoarc(filepath)


if __name__ == '__main__':
    config = conf()
    mail = mail(config)
    files = files(config)
    documents = documents(config, files)
    mail.findBills()
    files.extractallarchives()

    dirs = files.billdirs()
    for dir in dirs:
        bills = files.billfiles(dir)
        for bill in bills:
            billpath = files.billpath(bill, dir)
            if billpath:
                if billpath.lower().__contains__('.pdf'):
                    filetoprint = documents.makefiletoprint(billpath)
                    if filetoprint:
                        documents.makeagreement(filetoprint)
                        documents.printdocument(filetoprint)
                if os.path.basename(billpath).split('.')[-1].lower() in documents.WordDocTypes():
                    bill = documents.WordToPdf(billpath)
                    if bill:
                        filetoprint = documents.makefiletoprint(bill)
                        documents.makeagreement(filetoprint)
                        documents.printdocument(filetoprint)
