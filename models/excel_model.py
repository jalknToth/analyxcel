import pandas as pd
import os

class ExcelModel:
    @staticmethod
    def analyze_excel(filepath):
        
        try:
            # Read the Excel file
            df = pd.read_excel(filepath)
            
            # Prepare stats dictionary
            stats = {
                'filename': os.path.basename(filepath),
                'total_rows': len(df),
                'total_value': df.iloc[:, -1].sum(),
                'average_value': df.iloc[:, -1].mean(),
                'null_values': df.isnull().sum().sum(),
                'columns': list(df.columns),
                'data': df.to_dict('records')
            }
            
            return stats
        except Exception as e:
            raise ValueError(f"Error analyzing Excel file: {str(e)}")