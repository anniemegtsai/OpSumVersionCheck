from typing import List
import pandas as pd


class Sheet:
    def __init__(self, name, data_frame: pd.DataFrame, start_point: List = [1, 'A']):
        self.name = name
        self.start_point = start_point
        self.data_frame = data_frame
