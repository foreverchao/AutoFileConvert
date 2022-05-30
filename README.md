## CsvToExcel

## Push to GitLab
```sh
#### 不須上傳
1. 創建的虛擬機
2. .elasticbeanstalk
#### 上傳前
1. 需先退出虛擬機
退出虛擬機指令
$ deactivate.bat

```

## Redmine Issues 
```sh
https://cros-redmine.cros.tw/issues/4971

```

## Init
```sh
#### 1. Virtual Environment Init
python 3.7 或以上
virtualenv 20.10.0 或以上
pip 21.3.1 或以上

#### 2. Environment Init
folder 下創建虛擬機&進入
$ virtualenv virt
$ cd virt\Script\
$  activate.bat

回到folder下 安裝相關套件
$ pip install -r requirements.txt

```
## Run Dev
```sh
$ python application.py

```

## Deploy
```sh
虛擬機 專案 下 執行指令
$ pip freeze > requirements.txt

install the EB CLI
$ pip install awsebcli --upgrade

部署初始化

python 版本: python-3.7
應用程式名稱: flask-convert
AWS 地區: us-east-1
$ eb init -p python-3.7 flask-convert --region us-east-1


設置AWS金鑰(第一次建立才需要)
金鑰檔案URL:https://drive.google.com/file/d/1s5Isg1q33L1S08xuDVh7KjGP9aVxypcj/view?usp=sharing
You have not yet set up your credentials or your credentials are incorrect.
You must provide your credentials.
(aws-access-id): xxxxxxx
(aws-secret-key): xxxxxxx

$ eb init
C:\CHAO\issue4971\autofileconverttool (master -> origin)
(virt) λ eb init
Do you wish to continue with CodeCommit? (Y/n): n
Do you want to set up SSH for your instances?
(Y/n): y

Select a keypair.
(default is 21): 8


C:\CHAO\issue4971\autofileconverttool (master -> origin)


第一次創建
$ eb create flask-env

更新後續版本
$ eb deploy

開啟網頁
$ eb open

```





