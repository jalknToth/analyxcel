from flask import render_template

class ExcelView:
    @staticmethod
    def render_analysis(stats=None, columns=None, data=None, error=None):
        return render_template('analysis.html', 
                               stats=stats, 
                               columns=columns, 
                               data=data, 
                               error=error)