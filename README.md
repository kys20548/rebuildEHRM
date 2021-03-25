# rebuildEHRM

_將之前開發的小程式改寫成spring boot_

## 開發環境

| 使用技術 | 版本OR工具 |
| ------------- |:-------------:|
| java version|1.8|
| version control|github|
| DB|H2|
| 開發工具| IDEA|
| 程式碼打包建置|MAVEN|

## java 技術
* logback spring建議採用
* spring data jpa + lombok 底層簡單撰寫CRUD
* spring security 簡單登入
* AOP 實作audilog機制方便
* json fastjson號稱最快
* thymeleaf 取代JSP

## 撰寫風格

_盡量使用lambda方式撰寫_

_時間處理用 java8的新API_

## 預計實現功能
*專案能夠吃到外部(jar包外面)的設定檔案，例如txt，方便user不用動手改code，但可能還是要重啟專案，可能做個頁面顯示CONFIG目前的值*

*maven包程式時,可以透過指令選擇profile,做到不用改程式就能夠上版*

*寫個簡單TEST,測試model完成後,CRUD是否正常*
