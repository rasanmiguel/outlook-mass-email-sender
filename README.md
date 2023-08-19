# outlook-mass-email-sender
Simple outlook mass email sender that allows attachments.

What my code does:

Microsoft Outlook has a feature called 'Mail Merge' which enables you to send multiple personalized messages to multiple recipients.
This is great for sending multiple almost-similar emails that needs only minor changes in the mail body to suit each recipient (such as the Name and other details per email).
However, one thing that the feature is lacking is the inability to add attachments.

While there are third party add-ons for Outlook to enable the feature to add attachments, some companies have strict policies in installing such third party add-ons.
In addition, working with the Mail Merge tool is quite cumbersome. You have to go through a lot of configurations, and use both MS Word and MS Excel, in order to send the mails.

So I made a simple Python script that does exactly what Mail Merge does, with the addition of being able to add attachments.
It also has the ability to further improve through customization, instead of being limited to only configuration - which is a limiting factor in the case of Mail Merge.

For now, the script is only usable in the context of my work. I made the script to help my team send mass personzliaed emails with attachments.
In the future, I plan to:

1. Modify the script to take its data from an excel file. This will make it easier to consolidate the data.
2. Add proper error handling, and add an indicator to track which recipients have been successfully sent a mail
3. Convert script into a simple executable file, with a simple UI, making it more intuitive for non-technical users.
4. Adapt my code for broader applicability beyond the current context of my job.
5. Add font control and other stuff.
6. Optimize optimize optimize.

**************************************************************************************************************************************************************

How my code works:
1. Using pywin32 library, open an instance of Outlook (Note: It does not open the app, app does not need to be open for my code to work. It just creates an instance of Outlook in the background.)
2. I utilized a dictionary to reference the changes
For each item in the recipients list, perform the below actions:

	2.1 Set up the paramters for an individual mail
   
        2.1.a Create a new mail
        2.1.b Add Recipients
        2.1.c Add Subject (Email title)
        2.1.d Add Email Body
        2.1.e Add Attachments
        
    2.2 Send mail
