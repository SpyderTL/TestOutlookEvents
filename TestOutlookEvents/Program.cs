using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace TestOutlookEvents
{
	class Program
	{
		static void Main(string[] args)
		{
			var service = new ExchangeService
			{
				Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
			};

			//service.AutodiscoverUrl("email@email.com", Redirect);

			var credentials = new WebCredentials("email@email.com", "password");
			service.Credentials = credentials;

			var subscription = service.SubscribeToStreamingNotifications(new FolderId[] { WellKnownFolderName.Inbox }, EventType.NewMail);
			var connection = new StreamingSubscriptionConnection(service, new StreamingSubscription[] { subscription }, 30);

			connection.OnNotificationEvent += Connection_OnNotificationEvent;

			connection.Open();

			//var response = service.BindToFolders(new FolderId[] { WellKnownFolderName.Inbox }, PropertySet.FirstClassProperties);

			//var results = service.FindItems(WellKnownFolderName.Inbox, new ItemView(int.MaxValue));

			Console.ReadLine();

			connection.Close();
		}

		private static void Connection_OnNotificationEvent(object sender, NotificationEventArgs args)
		{
			foreach (var e in args.Events)
			{
				switch (e.EventType)
				{
					case EventType.NewMail:
						Console.WriteLine("New Mail");
						break;
				}
			}
		}

		private static bool Redirect(string redirectionUrl)
		{
			return true;
		}
	}
}
