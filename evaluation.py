from normalizaiton import Normalization
import numpy as np

class Evaluation():
    def __init__(self,text_df,correct_df):
        self.text_df = text_df
        self.correct_df = correct_df
        self.normalization = Normalization()
        
    def eval1(self):
        #データフレームを念の為正規化
        text_df = self.text_df.astype(str).applymap(self.normarization.normalize_and_strip).replace('nan', np.nan).dropna(how='all')
        correct_df = self.correct_df.astype(str).applymap(self.normalization.normalize_and_strip).replace('nan', np.nan).dropna(how='all')
        
        text_df.columns = [self.normalization.normalize_and_strip(col) for col in text_df.columns]
        correct_df.columns = [self.normalization.normalize_and_strip(col) for col in correct_df.columns]

        text_df = text_df.reset_index(drop=True)
        correct_df = correct_df.reset_index(drop=True)
        print(text_df)

        #各配列のテキストの長さを取得
        length_arr = correct_df.applymap(len).to_numpy() #各セルの文字列の長さの和
        mask_arr = (text_df == correct_df).to_numpy() #文字列が一致する箇所のみTrueにするmask配列
        return np.sum(self.normalization.relief_x(length_arr[mask_arr]))/np.sum(self.normalization.relief_x(length_arr)),mask_arr #Trueの箇所のみ文字列の長さを数え上げ


    def eval2(self):
        #データフレームを念の為正規化
        text_df = self.text_df.astype(str).applymap(self.normalization.normalize_and_strip).replace('nan', np.nan).dropna(how='all')
        correct_df = self.correct_df.astype(str).applymap(self.normalization.normalize_and_strip).replace('nan', np.nan).dropna(how='all')
        
        text_df.columns = [self.normalization.normalize_and_strip(col) for col in text_df.columns]
        correct_df.columns = [self.normalization.normalize_and_strip(col) for col in correct_df.columns]

        text_df = text_df.reset_index(drop=True)
        correct_df = correct_df.reset_index(drop=True)
        
        length_arr = correct_df.applymap(len).to_numpy()
        
        text_arr = text_df.to_numpy()
        correct_arr = correct_df.to_numpy()
        print(text_arr)
        match_score = np.array([[self.normalization.compare_text(a, b) for a, b in zip(row1, row2)] for row1, row2 in zip(text_arr, correct_arr)])
        return np.sum(match_score)/np.sum(np.ones((match_score.shape[0],match_score.shape[1])))