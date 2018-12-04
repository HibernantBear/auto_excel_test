import os
import xlrd
import xlwt


class AutoExcelTest():
    def __init__(self, project_path):
        self.project_path = project_path
        self.index = dict()
        self.load_answers()
        self.load_template()

    def load_answers(self):
        """
        load answer excel file
        even different test paper types, the answers store in one excel file named "answer.xlsx" but different sheets
        """
        self.answers = xlrd.open_workbook(
            os.path.join(self.project_path, "answer.xlsx"))

    def test_papers(self):
        """
        load test paper excels in folder "test_paper"
        :return: a test paper generator
        """
        test_paper_folder = os.path.join(self.project_path, "test_paper")
        test_paper_filenames = os.listdir(r"./demo/test_paper")
        for test_paper_filename in test_paper_filenames:
            print(test_paper_filename)
            yield xlrd.open_workbook(os.path.join(test_paper_folder, test_paper_filename)).sheet_by_index(0)

    def load_template(self):
        """
        load template excel file named "template.xlsx"
        there are two sheets in this file
        one is "input" which define the template of answer and test paper
        the other is "output" which define the template of output excel
        """
        self.template = xlrd.open_workbook(
            os.path.join(self.project_path, "template.xlsx"))

    def is_key(self, i):
        """
        judge whether the [i,0] cell in template file is a "KEY" type
        """
        return self.template.sheet_by_name("input").row(i)[0].value[:2] == "K_"

    def is_question(self, i):
        """
        judge whether the [i,0] cell in template file is a "QUESTION" type
        """
        return self.template.sheet_by_name("input").row(i)[0].value[:2] == "Q_"

    def get_question_score(self, i):
        """
        read the score in template file
        """
        return self.template.sheet_by_name("input").row(i)[1].value

    def get_answer(self, test_paper):
        """
        match the corresponding answer in answers, in most case this feature used for different test paper in the same examination
        """
        # print(test_paper.name)
        for answer in self.answers.sheets():
            match = True
            for index in range(0, answer.nrows):
                row = answer.row(index)
                if row[1].value and self.is_key(index):
                    # print(row[1].value, test_paper.row(index)[1].value)
                    if row[1].value != test_paper.row(index)[1].value:
                        match = False
            if match:
                # print("match"+answer.name)
                return answer
        # return False

    def caculate_score(self, test_paper):
        score = 0
        answer = self.get_answer(test_paper)
        for i in range(0, answer.nrows):
            if self.is_question(i):
                if answer.row(i)[1].value == test_paper.row(i)[1].value:
                    score += self.get_question_score(i)
        return score

    def prase_output(self, score=0, test_paper=False):
        """
        prase a list as a row in excel for display and output
        prase result based on template file
        """
        template_input = self.template.sheet_by_name("input")
        template_output = self.template.sheet_by_name("output")
        result = list()
        if test_paper:
            for i in range(0, template_output.ncols):
                key = template_output.col(i)[0].value
                if key == "SCORE":  # handle special mark
                    result.append(score)
                    continue
                result.append(test_paper.row(self.index[i])[1].value)
        else:
            for i in range(0, template_output.ncols):
                key = template_output.col(i)[0].value
                if key == "SCORE":  # handle special mark
                    result.append(template_input.col(3)[0].value)
                    continue
                for _i in range(0, template_input.nrows):
                    # match the output mark in the input
                    if template_input.row(_i)[0].value == key:
                        # generate index for fasten caculate
                        self.index[i] = _i
                        result.append(template_input.row(
                            _i)[1].value)  # generate title
                        break
        return result

    def caculate(self):
        """
        main loop
        """
        result = list()
        result.append(self.prase_output())
        for test_paper in self.test_papers():
            result.append(self.prase_output(
                self.caculate_score(test_paper), test_paper))
        # self.save(result)
        return result

    def save(self, result):
        """
        save the file to "output.xls"
        """
        output = xlwt.Workbook()
        sheet1 = output.add_sheet(u"sheet1", True)
        for r in range(0, len(result)):
            row = result[r]
            for l in range(0, len(row)):
                # print(r, l, result[r][l])
                sheet1.write(r, l, result[r][l])
        output.save(os.path.join(self.project_path, "output.xls"))


if __name__ == "__main__":
    print("Hello world!")
    aet = AutoExcelTest(r"./demo")
    result = aet.caculate()
    print(result)
    aet.save(result)
