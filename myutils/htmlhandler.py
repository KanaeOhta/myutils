from bs4 import BeautifulSoup, SoupStrainer


STRAINER = ["portal-header", "portal-body", "portal-footer"]


class HtmlHandler:

    def __init__(self, html_file):
        self.file = html_file


    def read(self, html_file=None):
        html_file = self.file if html_file is None else html_file
        with open(html_file, 'r', encoding='utf-8') as f:
            return f.read()


    def _parse(self, strainer, html_file=None):
        only_tags = SoupStrainer(**strainer)
        bs =  BeautifulSoup(
            self.read(html_file), 
            'lxml', 
            parse_only=only_tags
        )
        return bs


    def parse(self, *args):
        raise NotImplementedError()

