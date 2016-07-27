# outlook
automate outlook tasks
to use this file
1. make sure Developer Tab is added to outlook 2010; you can do this by going to options via Files Tab-> customize ribbon->check developer in MainTabs on right
2. press Alt+F11 on an open outlook window
3. Once VBA for outlook opens up, insert a new module and copy paste the code with changes to fields strPath and .Pattern
4. If you want to test it on current mail, remove arguments from Sub definition, declare and set an outlook mailitem that'll consume the current mail item
