# User-Termination-Automation-with-Powershell
![image](https://user-images.githubusercontent.com/91026684/134979337-d39be9f7-b492-4b27-b40d-50627cf2bbee.png)

I made this Powershell script in order to automate and standardize how you manage the termination of an employee/user.

>The Process requests username, email and if needed a ticket/process identification number.
>> Disable the account in AD.
>>> Force the Sync with Office365 to disable the account in azure too.
>>>> Revokes all devices/active sessions.
>>>>> Converts the Mailbox into a Shared Mailbox (This is recommended, because you don't need an 0365 license to keep those emails/information).
>>>>>> Gives you the chance to setup and Auto-reply message for the user.
>>>>>>> Gives you the chance to provide another user access to the previously created Shared Mailbox.
>>>>>>>> Remove user license from O365.
>>>>>>>>> The final log gets open with the confirmation of complection (Or failure) for every process.
