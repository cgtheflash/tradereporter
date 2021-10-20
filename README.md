# tradereporter
Summary:
Post trades made on the TradeStation platform to discord channel and create trading reports for trade tracking. 

Details:
1. Open local Outlook client as COM object
2. Fetch new unread emails from TradeStation in a folder named 'Trades'
3. Parse the email and create an object
4. Mark the email as read
5. Post the trade in a discord channel (webhook uri)
6. Track all daily trades and export to a daily csv
7. Track all weekly trades and export to a weekly csv
8. Track all monthly trades and export to a monthly csv
