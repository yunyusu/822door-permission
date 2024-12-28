# 822door-permission

The key is in my local desktop.

## plan
822門禁自動化表單

基本：google form提交後，python偵測到google sheet的變化後從中抓取資料(姓名、卡號)登入門禁系統，卡號登入成功後再傳送門禁申請已通過的訊息給使用者

進階：
1. 看準時間，當借用開始時間介於現在時間到借用時間結束前，再為使用者登入門禁系統。借用時間到時，拔除使用權限。
2. 另外列出現在正在使用該空間的人員
3. 增加須排除其他人員的大型借用需求選項(提前三天申請)，其餘進門前申請即可，需比對現在是否有大型活動借用中

## ref
1. pyautogui:
https://www.cc.ntu.edu.tw/chinese/epaper/home/20220920_006203.html
2. app script
https://www.youtube.com/watch?v=fKTnjZfPVh4&t=954s
