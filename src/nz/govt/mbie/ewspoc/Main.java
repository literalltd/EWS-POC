/**
 * 
 */
package nz.govt.mbie.ewspoc;

/**
 * @author Damian Rosewarne - MBIE
 *
 */

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.autodiscover.*;

class Main {
	static String email_address = "damian@literalgroup.com";
	static String email_password = "ewstesting123!";
	

	
	public Main() throws Exception {
	}

	/**
	 * @throws Exception 
	 */
	@SuppressWarnings("unused")
	
	public static void main(String[] args) throws Exception {

		/*
		 * Establish service and correct URL based on email address
		 */
		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(email_address, email_password);
		service.setCredentials(credentials);
		service.autodiscoverUrl(email_address, new RedirectionUrlCallback());

		/*
		 * Send Email using impersonated exchange account
		*/
		EmailMessage msg= new EmailMessage(service);
		msg.setSubject("Hello world!");
		msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Java API."));
		msg.getToRecipients().add("damian.rosewarne@mbie.govt.nz");
		//msg.send();
		
		
		/*
		 * Loop around Inbox associated with impersonated account and show message id, subject and body
		 */
		
		ItemView view = new ItemView(1000);
		FindItemsResults<Item> findResults = service.findItems(WellKnownFolderName.Inbox, view);
	        service.loadPropertiesForItems(findResults, PropertySet.FirstClassProperties);
		    for (Item item : findResults.getItems()) {
			System.out.println("id: " + item.getId());
			System.out.println("subject: " + item.getSubject());
			System.out.println("isNew?: " + item.isNew());
			System.out.println("hasAttachments?: " + item.getHasAttachments());
			System.out.println("body: " + item.getBody());
			System.out.println("==================================");
		}
		service.close();
	}

	
	static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        public boolean autodiscoverRedirectionUrlValidationCallback(
                String redirectionUrl) {
            return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }

}