import pandas as pd


class Sheet:
    def __init__(self, name, data_frame: pd.DataFrame):
        self.name = name
        self.data_frame = data_frame
