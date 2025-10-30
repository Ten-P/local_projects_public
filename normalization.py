import numpy as np
import unicodedata

class Normalization():
    @staticmethod
    def normalize_and_strip(text):
        if isinstance(text, str):
            normalized = unicodedata.normalize('NFKC', text)
            return normalized.replace('\u3000', ' ').replace('〜','~')
        return text
        
    def compare_text(self,original_text,correct_text):
        ori_text_l,corr_text_l = list(original_text),list(correct_text)
        count = 0
        for i in range(min(len(ori_text_l),len(correct_text))):
            if ori_text_l[i] == corr_text_l[i]:
                count += 1
        
        return count/min(len(ori_text_l),len(correct_text))

    #文字数が大きすぎる値に対して重みを加える（データの違いによるスコアの極端な増減を緩和）
    def relief_x(self,arr):
        return np.floor(np.log(arr))+1
        