# How to run
# python math2kor.py
# or
# python math2kor.py YourExelFile.xlsx

from openpyxl import load_workbook, Workbook
from TexSoup import TexSoup
import sys



class Eq2Script:
    def __init__(self):
        self.math_table = self.xlsx2dict('math_table.xlsx')

    # 엑셀을 딕셔너리로 변환
    def xlsx2dict(self, xlsx):
        load_wb = load_workbook(xlsx, data_only=True)
        load_ws = load_wb['Sheet1']

        math_table = {}

        for row in load_ws.rows:
            # 수식
            equation = row[0].value
            # 스크립트
            script = row[1].value

            math_table[equation] = script
        return math_table

    def convert_wrong_char(self, raw_text):
        return raw_text.replace('￦','\\')

    def post_process(self, text):
        text = text.replace('(',' ').replace(')','').replace('\\',' ').replace('.','점')
        return text

    def textree(self, node):    
        script = ''
        prev = None

        if node.name == 'frac':
            script += self.textree(TexSoup(node.args[1].value))
            script += '분의'
            script += self.textree(TexSoup(node.args[0].value))
        elif node.name != '[tex]':
            # \pm, \sqrt 등을 변환
            try:
                script += self.math_table.get(str(node.name))
            except:
                # 변환 테이블에 없을 경우 기호 이름 반환
                print(node.name)
            # \sqrt{...}에서 괄호 안에 있는 것을 변환
            for arg in node.args:
                script += self.textree(TexSoup(arg.value))
        else:
            for cont in node.contents:
                # 또 노드 타입이면 재귀 함수 부름
                if type(cont) == type(node):
                    script += self.textree(cont)
                # 텍스트면 그냥 출력
                else:
                    for c in cont:
                        # 특수 문자 인지 확인
                        if c in '^_':
                            prev = c
                        # 이전 문자가 특수 문자이면 함께 변환 (e.g. ^2 -> 제곱)
                        elif prev is not None:
                            script += self.math_table.get(prev + c)
                            prev = None
                        # 문자가 table에 있으면 변환
                        elif c in self.math_table:
                            script += self.math_table.get(c)
                        # table에 없으면 그대로 출력
                        else:
                            script += c
        return script

    def script(self, equation):
        node = TexSoup(equation)

        script = self.textree(node)
        return script

    def text2script(self, raw_text):
        text = ''
        for i, t in enumerate(self.convert_wrong_char(raw_text).split('$')):
            if (i % 2) == 0:
                text += t
            else:
                text += self.script(t)
        text = self.post_process(text)
        return text

    def xlsx2script(self, xlsx):
        load_wb = load_workbook(xlsx, data_only=True)
        load_ws = load_wb.active

        if load_ws.title + '_notex' in load_wb.sheetnames:
            save_ws = load_wb[load_ws.title + '_notex']
        else:
            save_ws = load_wb.create_sheet(load_ws.title + '_notex')

        for row_id, row in enumerate(load_ws.rows):
            for collumn_id, collumn in enumerate(row):
                raw_text = collumn.value
                if raw_text is not None: 
                    save_ws.cell(row_id+1, collumn_id+1, self.text2script(raw_text))

        load_wb.save(xlsx)
        print('저장완료...')

                
if __name__ == '__main__':
    tex_doc = r'$x=\frac{-b\pm\sqrt{b^{2}-4ac}}{2a}$'
    # tex_doc = r'$x^2+x^3\pm\frac{3x}{2}=4a-t$')

    sample = Eq2Script().text2script(tex_doc)
    print(sample)


    # Eq2Script().xlsx2script(sys.argv[1])

