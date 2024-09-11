<h1>Rebuild Outlook Autocomplete Address List</h1>


<h2>Description</h2>
The project consists of an AutoIt script that rebuilds the autocomplete cache based on unique addresses in the Sent Items folder in Outlook. The rebuild process can be time-consuming, depending on how many emails are in the Sent Items folder. The script compiles a list of unique email addresses from the Sent Items folder and creates an email message with those addresses as recipients. It then delays the sending of the email and subsequently deletes it. This process adds the email addresses to Outlook's autocomplete cache. The script notifies the user by displaying a 'Completed Rebuild' message.<br/>

<h2>Languages and Protocols</h2>

- <b>AutoIt</b>
- <b>Microsoft Outlook</b>

<h2>Environments Used </h2>

- <b>Outlook 2019 and older</b>

<h2>Program walk-through:</h2>

<p align="center">
Completed Rebuild: <br/>
<img src="https://i.imgur.com/2LRDdWw.png" height="80%" width="80%" alt="Completed Rebuild"/>
<br />
<br />
</p>

<!--
 ```diff
- text in red
+ text in green
! text in orange
# text in gray
@@ text in purple (and bold)@@
```
--!>
