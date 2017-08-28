# EWS-POC

Sample test of EWS Java API connecting to Office365 (Exchange 2010)

This repo contains a built ews-java-api JAR that was sourced from https://github.com/OfficeDev/

In order to run you need to update the main class with Office365 credentials:

  static String email_address = "xxxxxxxxxxxxx";
	static String email_password = "xxxxxxxxxxxxx";
  
In order to test the email sending capability, un remark the lines in the main class and add a email to send to:

  msg.getToRecipients().add("xxxxxxxxxxxxx");
