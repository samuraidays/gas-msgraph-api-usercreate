# gas-msgraph-api-usercreate
このスクリプトはスプレッドシートの情報を利用して、Office365にユーザを作成するスクリプトです。  

## 使用の想定イメージ
入社時のアカウント作成を想定しています。  

Googleフォームで入社する人に情報を入力してもらう  
→スプレッドシート内に情報がたまる  
→アカウント作成用のシートに自動で情報が作成される（セルの参照およびGASでローマ字への自動変換）  
→アカウント作成シートの情報を元にOffice365(MS Graph APIを使用)にアカウントを作成  
→作成されたアカウント情報がメール送信される  
→メールを印刷して入社する人に渡す

## スプレッドシート
![スプレッドシート](https://user-images.githubusercontent.com/4385484/75097610-fee22680-55ef-11ea-826b-537af9dcafdf.JPG)

## メール例
![メール](https://user-images.githubusercontent.com/4385484/75097725-72d0fe80-55f1-11ea-89e1-0c19eb371db6.JPG)
