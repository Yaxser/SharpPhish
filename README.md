# SharpPhish

This project was created to test an O365 module from an NDR vendor. The way it works is

1. Sends an email address via the outlook mailbox of the current user, with or without attachment. It will copy the current user signature, so it will look very legit.
2. Deletes the sent email from the "Sent" Folder
3. Waits for a reply to the email. If the reply arrives, it gets deleted before the user is notified.




What you will have to do to use it

1. Download, modify the source code to add subject, target, and content. [I will at some point add cli arguments; however, because we usually do assumed breached it is not a priority for me. Please make a PR if you have the time]
2. Compile
3. Move to target machine and run


What I want to do

1. Add error checking and report that to the operator
2. Add parameters
3. Add aggressor script

If you have suggestions or questions, feel free to reach out, [email](mailto:Y.Alhazmi@student.fontys.nl)
