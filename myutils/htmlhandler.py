from lxml import html


class HtmlHandler:

    def __init__(self, html_file, root):
        self.file = html_file
        self.root = root


    def to_string(self, tag):
        return html.tostring(tag, encoding='utf-8').decode()


    @classmethod
    def read_file(cls, html_file):
        with open(html_file, 'r', encoding='utf-8') as f:
            return f.read()


    @classmethod
    def get_root(cls, html_file):
        return html.fromstring(HtmlHandler.read_file(html_file))


    @classmethod
    def retrieve(cls, html_file, *args):
        root = HtmlHandler.get_root(html_file)
        return cls(html_file, root, *args)


