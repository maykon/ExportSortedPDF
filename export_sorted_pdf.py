# encoding: utf-8
import os
import io
import errno
import re
import glob
import json
import configparser
import logging
from PyPDF2 import PdfFileReader, PdfFileWriter

from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage


class ExportSortedPDF:

    def __init__(self):
        self.nomes = {}
        self.cfilename = "config.ini"
        self.config = configparser.ConfigParser()
        self.load_config()
        self.create_logger()
        self.config_log()
        self.nomes_file = os.path.join(self.tempdir, "nomes.txt")

    def create_logger(self):
        self.logger = logging.getLogger('ExportarPDF')
        logging.getLogger('pdfminer').setLevel(logging.ERROR)
        self.logger.setLevel(logging.INFO)

        ch = logging.StreamHandler()
        ch.setLevel(logging.ERROR)
        handler = logging.FileHandler('export.log', encoding = "UTF-8")
        handler.setLevel(logging.INFO)
        if self.isDebug():
          handler.setLevel(logging.DEBUG)

        # create a logging format
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        
        # add the handlers to the logger
        self.logger.addHandler(handler)
        self.logger.addHandler(ch)

    def create_config(self):
      cfgfile = open(self.cfilename,'w')
      try:        
        self.config.add_section('EXPORT')
        self.config.set('EXPORT', 'filename', 'sample.pdf')
        self.config.set('EXPORT', 'outputfile', 'sample_ordered.pdf')
        self.config.set('EXPORT', 'tempdir', './export')
        self.config.set('EXPORT', 'searchterm', 'Prezado\(a\)\s+\(cid:13\)([\w\s]+)')
        self.config.set('EXPORT', 'outputname', 'export\{}_{}.pdf')
        self.config.set('EXPORT', 'mergeterm', 'export\*.pdf')
        self.config.set('EXPORT', 'debug', 'False')
        self.config.set('EXPORT', 'deletetempdir', 'True')
        self.config.set('EXPORT', 'printtext', 'False')
        self.config.write(cfgfile)
      finally:
        cfgfile.close()
    
    def load_config(self):        
        if not os.path.isfile(self.cfilename):
          self.create_config()
        
        if not os.path.isfile(self.cfilename):
          raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.cfilename)

        self.config.read(self.cfilename)
        self.filename = str(self.config["EXPORT"]["filename"])
        self.outputfile = str(self.config["EXPORT"]["outputfile"])
        self.tempdir = str(self.config["EXPORT"]["tempdir"])
        self.searchterm = str(self.config["EXPORT"]["searchterm"])
        self.outputname = str(self.config["EXPORT"]["outputname"])
        self.mergeterm = str(self.config["EXPORT"]["mergeterm"])
        self.debug = self.config["EXPORT"].getboolean("debug", False)
        self.deletetempdir = self.config["EXPORT"].getboolean("deletetempdir", False)
        self.printtext = self.config["EXPORT"].getboolean("printtext", False)
    
    def config_log(self):
        if not self.isDebug():
          return

        self.logger.setLevel(logging.DEBUG)
        for key in self.config["EXPORT"]:
            self.logger.debug("Config %s: %s", key, self.config["EXPORT"][key])

    def isDebug(self):
      return self.debug

    def extract_text_by_page(self):
        if not os.path.isfile(self.filename):
          raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.filename)

        self.logger.info("Extraindo dados do arquivo %s", self.filename)
        with open(self.filename, 'rb') as fh:
            npage = 0
            for page in PDFPage.get_pages(fh,
                                          caching=True,
                                          check_extractable=True):
                resource_manager = PDFResourceManager()
                fake_file_handle = io.StringIO()
                converter = TextConverter(resource_manager, fake_file_handle)
                page_interpreter = PDFPageInterpreter(
                    resource_manager, converter)
                page_interpreter.process_page(page)

                text = fake_file_handle.getvalue()
                if self.printtext:
                  self.logger.debug(text)

                regex_search = re.search(self.searchterm, text)
                name = "page_{}".format(npage)
                if regex_search:
                  name = regex_search.group(1).strip()

                self.logger.debug("Termo: %s - Página: %s", name, npage)

                self.nomes[npage] = name
                npage += 1

                # close open handles
                converter.close()
                fake_file_handle.close()
          
        self.logger.info('Finalizando extração de dados')


    def criar_pdfs_termo(self):
        self.logger.info("Iniciando criação dos PDFs por termo.")
        pdf = PdfFileReader(self.filename)
        npages = pdf.getNumPages()
        for npage in range(npages):
            pdf_writer = PdfFileWriter()
            page = pdf.getPage(npage)
            pdf_writer.addPage(page)
            new_fname = self.nomes[npage]
            output_filename = self.outputname.format(
                new_fname, npage)

            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)

            self.logger.debug("Criado o arquivo: %s", output_filename)
        
        self.logger.info("Finalizado a criação dos PDFs por termo.")

    def remover_pdfs(self):
        self.logger.info('Removendo PDFs da pasta temporária.')
        dir = os.listdir(self.tempdir)
        for file in dir:
            if file.endswith(".pdf"):
                os.remove(os.path.join(self.tempdir, file))
    
    def remover_nomes_file(self):
      self.logger.info('Removendo arquivo nomes.txt.')
      if os.path.isfile(self.nomes_file):
        os.remove(self.nomes_file)  
      

    def limpar_ambiente(self):
        self.logger.info('Criando pasta temporária.')
        if not os.path.isdir(self.tempdir):
          os.makedirs(self.tempdir)

        self.remover_pdfs()
        self.remover_nomes_file()       

        self.logger.info('Removendo arquivo PDF ordenado.')
        if os.path.isfile(self.outputfile):
          os.remove(self.outputfile)

    def remover_temp_dir(self):
        if not self.deletetempdir:
          return
        
        self.logger.info('Apagando pasta temporária e arquivos.')
        self.remover_pdfs()
        self.remover_nomes_file()
        if os.path.isdir(self.tempdir):
          os.rmdir(self.tempdir)        

    def merge_pdf_files(self):
        self.logger.info("Iniciando a geração do PDF ordenado.")
        paths = glob.glob(self.mergeterm)
        paths.sort()
        pdf_writer = PdfFileWriter()

        for path in paths:
            pdf_reader = PdfFileReader(path)
            for page in range(pdf_reader.getNumPages()):
                pdf_writer.addPage(pdf_reader.getPage(page))

        with open(self.outputfile, 'wb') as fh:
            pdf_writer.write(fh)

        self.logger.info("Finalizando a geração do PDF ordenado.")

    def imprimir_lista_nomes(self):
        if not self.isDebug():
          return
        
        self.logger.debug("Quantidade de termos: %s", len(self.nomes))
        self.logger.debug("Gerando arquivo de termos 'termos.txt'")
        file = open(self.nomes_file, "w")
        for key, value in sorted(self.nomes.items(), key=lambda x: x[1]):
            file.write(value)
            file.write("\n")

        file.close()
        self.logger.debug("Finalizado a geração do arquivo.")

    def run(self):
        self.logger.info('Iniciando a exportação do PDF ordenado.')
        self.limpar_ambiente()
        self.extract_text_by_page()
        self.imprimir_lista_nomes()
        self.criar_pdfs_termo()
        self.merge_pdf_files()
        self.remover_temp_dir()
        self.logger.info('Finalizando rotina de exportação.')


if __name__ == '__main__':
    exp = ExportSortedPDF()
    exp.run()
