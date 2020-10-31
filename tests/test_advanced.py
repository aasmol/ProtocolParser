# -*- coding: utf-8 -*-

from .context import dataextractor

import docx


class TestParsing(object):

    def test_file_list(self):
        test_data = dataextractor.ProtocolData('c:/Users/user/Documents/Word parsing/Test docs')
        file_list = test_data.get_file_list()
        print(file_list)
        assert True

    def test_save(self):
        test_data = dataextractor.ProtocolData('c:/Users/user/Documents/Word parsing/Test docs')
        print(test_data.get_file_list())
        test_data.write_to_xlsx('c:/Users/user/Documents/Word parsing/parsed_data.xlsx')
        print(test_data.get_latchup_data())
        assert True

