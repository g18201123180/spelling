import os
import time
import tkinter as tk
from tkinter import messagebox as msgbox

import openai
import pandas as pd
import requests
from playsound import playsound
from tqdm import tqdm

from utils.AuthV3Util import addAuthParams

# 单词初始计数
init_wrong = 3
# 拼写错误增加计数
punish = 3
# 每日单词数量
daily_num = 50

# 网易有道tts
# 您的应用ID
APP_KEY = 
# 您的应用密钥
APP_SECRET = 

# open ai key
client = openai.OpenAI(api_key="")

# 固定提示词
base_prompt = """翻译我给你的单词，内容之间不要空行，不要有字体加粗，不要返回翻译以外的内容。可选词性有：n.,v.,vt.,vi.,adj.,adv.,prep.,conj.,pron.,art.,num.,int.,aux.,modal v.,det.,phr。格式如下(替换[]内容，不保留[])：[词性]\n1. [释义1] ：[例句1]\n [释义2] ：[例句2]\n...\n[词性]\n...\n"""

path = project_dir = os.path.dirname(os.path.abspath(__file__))

def translate_word(word):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": base_prompt},
            {"role": "user", "content": word}
        ]
    )
    return response.choices[0].message.content

def word2blank(word):
    i = 0
    f = 1
    ans = ''
    while i < len(word) - 1:
        if f == 1:
            ans += word[i] + ' '
            f = 0
        elif word[i] == ' ':
            ans += '  '
            f = 1
        else:
            ans += '_ '
        i += 1
    ans += '_'
    return ans


def createRequest(word,path):
    '''
    note: 将下列变量替换为需要请求的参数
    '''
    q = word
    voiceName = 'youxiaomei'
    format = 'mp3'

    data = {'q': q, 'voiceName': voiceName, 'format': format}

    addAuthParams(APP_KEY, APP_SECRET, data)

    header = {'Content-Type': 'application/x-www-form-urlencoded'}
    res = doCall('https://openapi.youdao.com/ttsapi', header, data, 'post')
    saveFile(res,path)


def doCall(url, header, params, method):
    if 'get' == method:
        return requests.get(url, params)
    elif 'post' == method:
        return requests.post(url, params, header)
    return None


def saveFile(res,path):
    contentType = res.headers['Content-Type']
    if 'audio' in contentType:
        fo = open(path, 'wb')
        fo.write(res.content)
        fo.close()
    else:
        print(str(res.content, 'utf-8'))


class Spelling:
    def __init__(self, master):
        self.remember = 0
        self.index_df = None
        self.index = 0
        self.pageList = []
        self.pageIndex = 0
        self.dataIndex = -1
        self.data = {}

        # 下载单词
        self.get_words()

        self.master = master
        self.master.title("单词拼写")
        self.master.protocol("WM_DELETE_WINDOW", self.close)

        # 显示器
        self.indicator = tk.StringVar()
        self.indicator_label = tk.Label(self.master, textvariable=self.indicator, font=("Arial", 20), width=50,
                                        height=1, anchor="center")
        self.indicator_label.grid(row=0, column=0)

        self.word = tk.StringVar()
        self.word_label = tk.Label(self.master, textvariable=self.word, font=("Arial", 24), width=50, height=3,
                                   anchor="center")
        self.word_label.grid(row=1, column=0)

        self.paraphrase = tk.StringVar()
        width=70
        self.paraphrase_label = tk.Label(self.master, textvariable=self.paraphrase, font=("Microsoft YaHei", 12),
                                         width=width, height=10, anchor="nw", justify="left",wraplength=width * 15)
        self.paraphrase_label.grid(row=2, column=0)

        self.entry = tk.Entry(self.master, width=50, font=("Arial", 18))
        self.entry.grid(row=3, column=0, padx=10, pady=30, ipadx=10, ipady=10)
        self.entry.bind("<Return>", self.verify)

        self.next_word()

    def get_words(self):
        self.index_df = pd.read_excel(f"{path}\\index.xlsx")

        total_unreviewed = 0
        for idx, row in self.index_df.iterrows():
            df_id=int(row['id'])
            filename = f"{path}\\word lists\\{df_id}.xlsx"

            if row['download'] == 0:
                print(f"下载{filename}中：")
                self.get_words_of_excel(filename)
                self.index_df.loc[idx, 'download'] = 1

            df = pd.read_excel(filename)
            count = (df['wrong'] != 0).sum()
            if count != 0:
                self.pageList.append(df_id)
                self.data[df_id] = df

            total_unreviewed += count
            if total_unreviewed >= daily_num:
                break
        self.index_df.to_excel(f"{path}\\index.xlsx", index=False)

    def get_words_of_excel(self, excel_path):
        df = pd.read_excel(excel_path, sheet_name='Sheet1')
        df['paraphrase'] = df['paraphrase'].astype('object')
        for i in tqdm(range(len(df))):
            df.loc[i, 'wrong'] = init_wrong
            sound_path = f"{path}\\sound\\{df['word'][i]}.mp3"
            sound_path = sound_path.replace(' ', '_')
            createRequest(df['word'][i],sound_path)
            df.loc[i, 'paraphrase'] = translate_word(df['word'][i])
            time.sleep(0.75)


        df["wrong"] = df.groupby("word")["wrong"].transform(lambda x: x.sum())
        df = df.drop_duplicates('word', keep='first').reset_index()
        df.drop(['index'], axis=1, inplace=True)
        self.save(df, excel_path)

    def verify(self, event):
        self.word.set(self.data[self.pageList[self.pageIndex]]['word'][self.dataIndex])
        self.paraphrase.set(self.data[self.pageList[self.pageIndex]]['paraphrase'][self.dataIndex])
        if self.entry.get() == self.data[self.pageList[self.pageIndex]]['word'][self.dataIndex]:
            playsound(f'{path}\\sound\\0.wav')
            self.data[self.pageList[self.pageIndex]].loc[self.dataIndex, 'wrong'] -= 1
            if self.data[self.pageList[self.pageIndex]].loc[self.dataIndex, 'wrong']==0:
                self.remember+=1
            self.master.after(1000, self.next_word)
        else:
            playsound(f'{path}\\sound\\1.wav')
            self.data[self.pageList[self.pageIndex]].loc[self.dataIndex, 'wrong'] += punish
            self.word_label['fg'] = 'red'
            self.master.after(3000, self.repeatWord)
        self.entry.delete(0, tk.END)

    def find_next_wrong_index(self):
        for idx in range(self.dataIndex+1, len(self.data[self.pageList[self.pageIndex]])):
            if self.data[self.pageList[self.pageIndex]].loc[idx, 'wrong'] != 0:
                self.dataIndex=idx
                return True
        if self.pageIndex>=len(self.pageList):
            return False
        self.pageIndex+=1
        self.dataIndex=-1
        for idx in range(self.dataIndex + 1, len(self.data[self.pageList[self.pageIndex]])):
            if self.data[self.pageList[self.pageIndex]].loc[idx, 'wrong'] != 0:
                self.dataIndex=idx
                return True
        return False

    def next_word(self):
        self.index += 1
        self.isFinish()

        self.find_next_wrong_index()

        self.word_label['fg'] = 'black'
        word = self.data[self.pageList[self.pageIndex]]['word'][self.dataIndex]
        sound_path = f"{path}\\sound\\{self.data[self.pageList[self.pageIndex]]['word'][self.dataIndex]}.mp3"
        sound_path = sound_path.replace(' ', '_')
        playsound(sound_path)

        self.isFinish()
        self.indicator.set(str(self.index) + ' / ' + str(daily_num))
        blank_word=word2blank(word)
        self.word.set(blank_word)
        self.paraphrase.set(self.data[self.pageList[self.pageIndex]]['paraphrase'][self.dataIndex].replace(word, blank_word))

    def repeatWord(self):
        self.word_label['fg'] = 'black'
        word = self.data[self.pageList[self.pageIndex]]['word'][self.dataIndex]
        self.word.set(word2blank(word))

    def save_all(self):
        self.index_df.to_excel(f"{path}\\index.xlsx", index=False)
        for id_, df in self.data.items():
            file_path = f"{path}\\word lists\\{id_}.xlsx"
            df.to_excel(file_path, index=False)

    def save(self, df, excel_path):
        with pd.ExcelWriter(excel_path) as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)

    def close(self):
        self.save_all()
        self.master.destroy()

    def isFinish(self):
        if self.index > daily_num:
            self.close()
            msgbox.showinfo("提示", f"今日目标已达成，共记忆{self.remember}个单词")
            exit()


if __name__ == "__main__":
    root = tk.Tk()
    Spelling(root)
    root.mainloop()
