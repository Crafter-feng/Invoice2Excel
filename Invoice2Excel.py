# -*- coding:utf-8 -*-

"""
parse PDF invoice and extract data to Excel
"""
import pdfplumber as pb
import pandas as pd
import os
import re
import sys
import getopt

class Extractor(object):
    def __init__(self, path):
        self.file = path

    @staticmethod
    def load_files(directory):
        """load files"""
        paths = []
        for file in os.walk(directory):
            for f in file[2]:
                path = os.path.join(file[0], f)
                if os.path.isfile(path) and os.path.splitext(path)[1] == '.pdf':
                    paths.append(path)
        return paths

    def _strip_words(self, words):
        table = {ord(f):ord(t) for f,t in zip(
             u'、，。：！？【】￥（）％＃＠＆１２３４５６７８９０',
             u',,.:!?[]¥()%#@&1234567890')}
        if isinstance(words,list):
            words = [self._strip_words(i) for i in words]   
        elif isinstance(words,dict):
            for k,v in words.items():
                words[k] = self._strip_words(v)
        elif isinstance(words,str):
            words = re.sub("\u3000|\n|\xa0| |\t", '',words).translate(table)
        elif words is None:
            words = ''
        return words


    def _load_data(self):
        if self.file and os.path.splitext(self.file)[1] == '.pdf':
            pdf = pb.open(self.file)
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=5, keep_blank_chars=True)
            words = self._strip_words(words)
            image = page.to_image(resolution=150)          
            return {'words': words, 'image': image}
        else:
            print("file %s can't be opened." % self.file)
            return None

    def _extrace_from_words(self, words):
        """ 从单词中提取 """
        info = {}
        lines = {}
        hlines = {}
        for idx,word in enumerate(words):          
            if word['text'] == '' or word['text'] in '名合':
                continue  
            elif word['text'] in "销售方购买方密码区备注":
                pos = round(word['x0'])
                if hlines.get(pos):
                    hlines[pos].append(word)
                else:
                    hlines[pos] = [word]
            else: 
                top = round(word['top'])
                bottom = round(word['bottom'])
                pos = (top + bottom) // 2
                if lines.get(pos):
                    lines[pos].append(word)
                else:
                    lines[pos] = [word]

        hlines_pack = []
        last_pos = None
        for pos in sorted(hlines):
            arr = hlines[pos]

            if len(hlines_pack) > 0 and pos - last_pos <= 5:
                hlines_pack[-1] += arr
                continue

            hlines_pack.append(arr)
            last_pos = pos
        
        hinfo = {}   

        for line in hlines_pack:
            line.sort(key=lambda t: t['top'])
            t = "".join([ j['text'] for j in line])
            for tx in ["购买方", "销售方", "密码区", "备注"] :
                x = t.find(tx)
                if x >= 0:  
                    idx = x+len(tx)-1
                    if len(line) <= idx:
                        idx = -1
                    hinfo[tx] =  {'x0': line[x]['x0'],'x1': line[x]['x1'],'top': line[x]['top'],'bottom': line[idx]['bottom']}
        
        lines_pack = []
        last_pos = None
        for pos in sorted(lines):
            arr = lines[pos]

            if len(lines_pack) > 0 and pos - last_pos <= 5:
                lines_pack[-1] += arr
                continue

            lines_pack.append(arr)
            last_pos = pos
            
        for i,line in enumerate(lines_pack):
            lines_pack[i].sort(key=lambda t: t['x0'])
            
        isSeller = "购买方"
        _lines_pack = []
        for ldx,pack in enumerate(lines_pack):
            for idx, _line in enumerate(pack):
                line = _line["text"]
                if '电子普通发票' in line:
                    info['标题'] = line

                elif '发票代码:' in line:
                    info['发票代码'] = line.split(':')[1]

                elif '发票号码:' in line:
                    info['发票号码'] = line.split(':')[1]

                elif '开票日期:' in line:
                    info['开票日期'] = line.split(':')[1]

                elif '机器编号:' in line:
                    info['机器编号'] = line.split(':')[1]
                    if info['机器编号'] == '':
                        info['机器编号'] = pack[idx+1]['text']    
                    hinfo["校验码"] = _line
                    
                elif '校验码:' in line:
                    info['校验码'] = line.split(':')[1] 
                    hinfo["校验码"] = _line
                    
                elif '价税合计' in line:
                    isSeller = "销售方"
                    info['价税合计(大写)'] = pack[idx+1]['text']
                    info['价税合计(小写)'] = pack[-1]['text'].split('¥')[1]
                    
                elif '合计' in line or line == '计':
                    isSeller = "销售方"
                    info['合计(金额)'] = pack[idx+1]['text'].split('¥')[1]
                    if '¥' in pack[idx+2]['text']:
                        info['合计(税额)'] = pack[idx+2]['text'].split('¥')[1]
                    else:
                        info['合计(税额)'] = ""
                    
                elif '收款人:' in line:
                    info['收款人'] = line.split(':')[1]
                    if info['收款人'] == '':
                        info['收款人'] = pack[idx+1]['text']
                        
                elif '复核:' in line:
                    info['复核'] = line.split(':')[1]
                    if info['复核'] == '':
                        info['复核'] = pack[idx+1]['text']
                        
                elif '开票人:' in line:
                    info['开票人'] = line.split(':')[1]
                    if info['开票人'] == '':
                        info['开票人'] = pack[idx+1]['text']
                        
                elif "名称:" in line or line.startswith("称:"):
                    info[f'名称({isSeller})'] = line.split(':')[1]
                    
                elif "纳税人识别号:" in line or line.startswith("称:"):
                    info[f'纳税人识别号({isSeller})'] = line.split(':')[1]
                    
                elif "地址" in line and "电话" in line:
                    info[f'地址、电话({isSeller})'] = line.split(':')[1]
                    
                elif "开户行及账号:" in line:
                    info[f'开户行及账号({isSeller})'] = line.split(':')[1]
                else:
                    if "税率" in line:
                        hinfo["税率"] = _line
                    _lines_pack.append(_line)
        
        # 密码提取
        lines_pack = _lines_pack
        info["密码区"] = ""
        x1 = hinfo["密码区"]['x1']
        top = hinfo["校验码"]['bottom']
        bottom = hinfo["税率"]['top']
        
        for ldx,pack in enumerate(lines_pack):
            if pack['x0'] > x1 and pack['top'] > top and pack['bottom'] < bottom:
                info["密码区"] += pack['text']
            
        return info

    def extract(self):
        data = self._load_data()
        words = data['words']
        data['info'] = self._extrace_from_words(words)
        df = pd.DataFrame([data['info']])
        return df


if __name__ == '__main__':
    IN_PATH = 'example'
    OUT_PATH = 'result.xlsx'
    # parse params
    opts, args = getopt.getopt(sys.argv[1:], 'p:ts:', ['test', 'path=', 'save='])
    for opt, arg in opts:
        if opt in ['-p', '--path']:
            IN_PATH = arg
        elif opt in ['--test', '-t']:
            IN_PATH = 'example'
        elif opt in ['--save', '-s']:
            OUT_PATH = arg
    # run programme
    print(f'run {"test" if IN_PATH == "example" else "extracting"} mode, load data from directory {IN_PATH}.\n{"*"*50}')
    files_path = Extractor('').load_files(IN_PATH)
    num = len(files_path)
    print(f'total {num} file(s) to parse.\n{"*"*50}')
    data = pd.DataFrame()
    result = []
    for index, file_path in enumerate(files_path):
        print(f'{index+1}/{num}({round((index+1)/num*100, 2)}%)\t{file_path}')
        extractor = Extractor(file_path)
        try:
            d = extractor.extract()
            data = pd.concat([data, d], axis=0, sort=False, ignore_index=True)
            result.append(["succeed", file_path])
        except Exception as e:
            result.append(["fail", file_path])
            print('file error:', file_path, '\n', e)
            
    print(f'{"*"*50}\nfinish parsing, save data to {OUT_PATH}')

    result = pd.DataFrame(result, columns=["解析状态" , "文件路径"])
    writer = pd.ExcelWriter(OUT_PATH)
    data.to_excel(writer,'data')
    result.to_excel(writer,'result')
    writer.save()

    print(f'{"*" * 50}\nALL DONE. THANK YOU FOR USING MY PROGRAMME. GOODBYE!\n{"*"*50}')

