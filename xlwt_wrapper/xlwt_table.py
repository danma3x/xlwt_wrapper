import xlwt

class xlwt_table:

    def __init__(self, data=None, xlwt_sheet=None, starts_at=None):
        """
        This class is designed to let your write your table(assumed 2d array) with given offset and let you
        convert the indices to Excel string representations with intent of making formulas easier to write
        :param data:
        :param xlwt_sheet:
        :param starts_at:
        :return:
        """
        self.data = data
        self.xlwt_sheet = xlwt_sheet
        self.starts_at = starts_at
        if self.starts_at is not None:
            if type(self.starts_at) == type(str):
                self.starts_at = self.deduce_index(self.starts_at)
        else:
            self.starts_at = [0, 0]

    def set_data(self, data):
        self.data = data

    def set_sheet(self, xlwt_sheet):
        self.xlwt_sheet = xlwt_sheet

    def write_table(self):
        """
        write the table to the sheet considering the offset if the one was set
        :return:
        """
        for i, row in enumerate(self.data):
            for j, cell in enumerate(row):
                self.xlwt_sheet.write(self.starts_at[1] + i,self.starts_at[0] + j, cell)

    @staticmethod
    def deduce_index(text_index):
        """
        return index as an int array from the excel index for xlwt
        :param text_index:
        :return:
        """
        text_index = text_index.capitalize()
        return [ord(text_index[0]) - 65, int(text_index[1:]) - 1]

    def get_literal(self, index):
        """get string representation of the index(for excel formulas)"""
        return chr(self.starts_at[0] + index[0] + 65) + str(self.starts_at[1] + index[1] + 1)