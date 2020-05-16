from io import StringIO
from xml.etree import ElementTree as ET


class XMLHandler:

    def __init__(self, xml_file, root, namespaces):
        self.file = xml_file
        self.root = root
        self.namespaces = namespaces


    def normalize(self, attrib_dict):
        """Removes namespace from keys of attribute dict.
        """
        return {k.split('}')[1] if k.startswith('{') else \
            k : v for k, v in attrib_dict.items()}


    @classmethod    
    def read_file(cls, xml_file):
        with open(xml_file, 'r', encoding='utf-8') as f:
            return f.read()


    @classmethod
    def get_namespaces(cls, xml_file):
        """Gets namespaces and give key to default namespace.
        """
        namespaces = dict(
            [node for _, node in \
                ET.iterparse(StringIO(XMLHandler.read_file(xml_file)), events=['start-ns'])]
        )
        namespaces['uri'] = namespaces['']
        return namespaces


    @classmethod
    def get_root(cls, xml_file):
        return ET.fromstring(XMLHandler.read_file(xml_file))
    

    @classmethod
    def retrieve(cls, xml_file, *args):
        namespaces = XMLHandler.get_namespaces(xml_file)
        root = XMLHandler.get_root(xml_file)
        return cls(xml_file, root, namespaces, *args)


    def parse(self, *args):
        raise NotImplementedError()

