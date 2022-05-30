from flask import Flask
from crosToCIC import  csvToExcel

application = Flask(__name__)

application.config['JSON_AS_ASCII'] = False

# 路由

# 根目錄/
# 檢查server是否正常運作
@application.route('/',methods=["GET"])
def index():
    
    return "Health OK"

# /csvToCicExcel
# csv 檔案轉換 為 excel
@application.route('/csvToCicExcel',methods=["POST"])
def csvToCicExcel():
    resp = csvToExcel()
    return resp

@application.route('/test',methods=["GET"])
def test():
    
    return "test OK!"



if __name__ == '__main__':
    application.run(debug = True)