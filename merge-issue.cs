using System;
using System.Collections.Generic;
using System.Linq;
using SRK.Crm.Main.Common;
using SRK.Crm.Main.Common.Entities;
using SRK.Crm.Main.Common.DataContracts;
using SRK.Crm.Main.CrmServices.DataContracts;
using SystemJobInformation = SRK.Crm.Main.Common.SystemJobInformation;
using Microsoft.Xrm.Sdk;
using System.IO;
using System.Globalization;

namespace SRK.Crm.Main.CrmServices.CrmServiceFacade
{
    //[ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Multiple, InstanceContextMode = InstanceContextMode.PerSession)]
    public class CrmMainService : ICrmMainService
    {
        /// <summary>
        /// Metod för att ta emot alla typer av betalningar från biztalk (IF030)
        /// </summary>
        /// <param name="payment">Payment-objekt med uppgifter för betalningen.</param>
        public void ImportAllPaymentsInCRM(Payment payment)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera payment
                    if (payment == null)
                        throw new ArgumentNullException("payment");

                    // Validera att vi har ett CustomerId
                    if (string.IsNullOrWhiteSpace(payment.CustomerId))
                        throw new ArgumentException("CustomerId får ej vara blankt.", "payment.CustomerId");

                    // Validera att kampanj har specificerats för gåva
                    if (string.IsNullOrWhiteSpace(payment.CampaignId) && (payment.PaymentType == "G" || payment.PaymentType == "AG"))
                        throw new ArgumentException("CampaignId får ej vara blankt för gåvor.", "payment.CampaignId");

                    // Validera att verifikationsnummer från AX har specificerats
                    if (string.IsNullOrWhiteSpace(payment.VoucherNumber))
                        throw new ArgumentException("VoucherNumber får ej vara blankt.", "payment.VoucherNumber");

                    // Validera att transaktionsnr från AX har specificerats
                    if (string.IsNullOrWhiteSpace(payment.TransactionNumber))
                        throw new ArgumentException("TransactionNumber får ej vara blankt.", "payment.TransactionNumber");
                    //    payment.TransactionNumber = ""; //Transaktionsnummer ej obligatoriskt i nuläget

                    // OwnerType styr vem som gjort betalningen. 
                    // P=Private person, F=Company, O=Organization
                    if (payment.OwnerType != "P" && payment.OwnerType != "F" && payment.OwnerType != "O" && payment.OwnerType != "VM")
                        throw new ArgumentException("Felaktig OwnerType, giltiga värden P,F eller O.", "payment.OwnerType");

                    //TODO: Ska vi göra någon kontroll på om transaktionsnummret redan finns - så det inte är en betalning som skickas dubbelt?
                    if (dal.CheckIfTransactionNumberExists(payment.TransactionNumber, payment.PaymentType))
                    {
                        exception = new ArgumentException("En betalning med ett transaktionsnummer som redan finns i CRM skickades in. Import av betalningen avbröts.", "payment.TransactionNumber");
                        return;
                    }

                    switch (payment.PaymentType)
                    {
                        case "G": // Gåva
                            {
                                CreateDonationFromPayment(payment, dal);
                                break;
                            }
                        case "AG": // Autogirobetalning för gåva
                            {
                                CreateDonationFromAutogiroPayment(payment, dal);
                                srk_autogiro autogiro = UpdateAutogiroWithValidPayment(payment.PaymentNumber, payment.PaymentDate);

                                if (autogiro != null)
                                {
                                    SrkCrmContext.Instance.GetSubscriptions(dal).HandleNewMonthlyDonorPaymentForContact(autogiro);
                                }

                                break;
                            }
                        case "M": // Medlemsskapsbetalning
                            {
                                UpdateMembershipPaymentInCRM(payment, dal, false);
                                break;
                            }
                        case "AM": // Autogiro-medlemsskapsbetalning
                            {
                                UpdateMembershipPaymentInCRM(payment, dal, true);
                                break;
                            }
                        case "P": // Produkt
                            {
                                throw new Exception("Mottagen PaymentType = P. Produktköp stöds inte av systemet idag.");
                                //break;
                            }
                        case "UTB": // Utbildning
                            {
                                dal.UpdateCoursePaymentInCRM(payment.PaymentNumber, payment.PaymentDate);
                                break;
                            }

                        default:
                            throw new ArgumentException("Felaktigt värde för PaymentType. Giltiga värden AG, G, M, AM, P, VM, UTB.", "Payment.PaymentType");
                    }
                }
                catch (Exception ex)
                {
                    //Om ett existerande värvmingsmedlemskap hittas på inkommen betalning skall aktivitet skapas och inget fel skall slängas.
                    if (!ex.Message.Equals("Manual activity created"))
                    {
                        exception = ex;
                        throw;
                    }
                }
                finally
                {
                    try
                    {
                        SrkCrmContext.LogOperation(dal, "CrmMainService::ImportAllPaymentsInCRM",
                        ObjectToString.ToString(payment) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }



        private void SetMovemementOrganisation(Payment payment, CrmDataAccess dal)
        {
            var membership = dal.GetMembershipFromOCRNumber(payment.PaymentNumber);

            if (membership != null && membership.srk_contactId != null && membership.srk_organizationId != null)
            {
                var membershipContact = dal.Context.ContactSet.First(x => x.Id == membership.srk_contactId.Id);

                var membershipMaximumAgeForYouth = int.Parse(dal.GetConfiguration("MembershipMaximumAgeForYouth"));
                var age = CrmUtility.GetAgeOfContact(membershipContact);

                var isValidForOrganizationMember = (age <= membershipMaximumAgeForYouth);

                //om giltig för rörelsemedlemsskap 
                if (isValidForOrganizationMember)
                {
					// REF2016CHANTE 26
//					CrmEntityReference organisationReference = null;
					EntityReference organisationReference = null;

					// REF2016CHANTE 26
//					var selectedOrganization = dal.GetOrganization(new CrmEntityReference(srk_organization.EntityLogicalName, membership.srk_organizationId.Id));
					var selectedOrganization = dal.GetOrganization(new EntityReference(srk_organization.EntityLogicalName, membership.srk_organizationId.Id));
					if (selectedOrganization.srk_organizationtype.Value == (int)MemberType.MemberRKUF)
                    {
                        organisationReference = dal.GetOrganizationFromPostalCode(membershipContact.Address1_PostalCode, MemberType.MemberSRK.ToString()).ToEntityReference();
                    }
                    else
                    {
                        organisationReference = dal.GetOrganizationFromPostalCode(membershipContact.Address1_PostalCode, MemberType.MemberRKUF.ToString()).ToEntityReference();
                    }

                    membership.srk_IsOrganizationMember = true;
                    membership.srk_OrganizationMembership_OrgId = organisationReference;
                    dal.Update(membership);
                }
            }

        }

        public srk_membership CreateMembershipFromPayment(Payment payment, ICrmDataAccess dataAccess, bool yearIsClosed)
        {
			// REF2016CHANTE 26
//			CrmEntityReference movementOrganization = null;
			EntityReference movementOrganization = null;
			bool isMovementMembership = false;
            bool isYouth = false;
            int age = -1;

            var matchedContacts = dataAccess.Context.ContactSet.Where(x => x.srk_contactid == payment.CustomerId).ToList();

            Contact contact = null;
            if (matchedContacts.Count() == 1)
                contact = matchedContacts.Single();
            else
                throw new CrmRedyException(matchedContacts.Count() + " contacts found in CRMMainService.CreateMembershipFromPayment for: payment.CustomerId = " +
                                            (payment.CustomerId != null ? payment.CustomerId : "<null>" + Environment.NewLine + ObjectToString.ToString(payment)), true);


            // membership type
            var membershipTypeRKUF = dataAccess.Context.srk_membershiptypeSet.Where(x => x.srk_Internalname == SRK.Crm.Main.Common.MemberType.MemberRKUF.ToString()).Single();
            var membershipTypeSRK = dataAccess.Context.srk_membershiptypeSet.Where(x => x.srk_Internalname == SRK.Crm.Main.Common.MemberType.MemberSRK.ToString()).Single();

            int membershipMaximumAgeForYouth = 32;
            try
            {
                membershipMaximumAgeForYouth = int.Parse(dataAccess.GetConfiguration("MembershipMaximumAgeForYouth"));
            }
            catch (Exception)
            {
                throw new CrmRedyException("Konfigurationen för åldersgräns för RKUF kunde inte hämtas.");
            }

            if (string.IsNullOrWhiteSpace(contact.srk_socialsecuritynumber) == false)
            {
                age = CrmUtility.GetAgeOfContact(contact);
                isYouth = (age <= membershipMaximumAgeForYouth);
            }

			//Kontrollera om kontakten är över 31 år och betalt <= 50kr. Då skall en aktivitet skapas för manuell hantering.
			//REF2016CHANTE 14
//			if (!isYouth && payment.Amount <= membershipTypeRKUF.srk_price && age != -1)
            if (!isYouth && payment.Amount <= membershipTypeRKUF.srk_price?.Value && age != -1)
			{
				SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                    "Individen är över 31 år och har betalt mindre än 50kr",
                    "En betalning inkom på en individ som är äldre än 31 år, dock är inbetalt belopp mindre än 50kr. Individ id: " + payment.CustomerId + Environment.NewLine + Environment.NewLine + ObjectToString.ToString(payment),
                    contact);

                throw new Exception("Manual activity created");
            }

            srk_membershiptype membershipType = null;
            MemberType membershipTypeEnum = MemberType.MemberSRK;
			//REF2016CHANTE 14
//			if (payment.Amount > membershipTypeRKUF.srk_price)
            if (payment.Amount > membershipTypeRKUF.srk_price?.Value)
			{
				membershipType = membershipTypeSRK;
                membershipTypeEnum = MemberType.MemberSRK;

            }
            else
            {
                membershipType = membershipTypeRKUF;
                membershipTypeEnum = MemberType.MemberRKUF;
            }

			//REF2016CHANTE 14
//			if (isYouth && payment.Amount > membershipTypeRKUF.srk_price)
            if (isYouth && payment.Amount > membershipTypeRKUF.srk_price?.Value)
			{
				movementOrganization = dataAccess.GetOrganizationRefFromPostalCode(contact.Address1_PostalCode, MemberType.MemberRKUF.ToString());
                isMovementMembership = true;

            }
			//REF2016CHANTE 14
//			else if ((isYouth && payment.Amount <= membershipTypeRKUF.srk_price) || (age == -1 && payment.Amount <= membershipTypeRKUF.srk_price))
            else if ((isYouth && payment.Amount <= membershipTypeRKUF.srk_price?.Value) || (age == -1 && payment.Amount <= membershipTypeRKUF.srk_price?.Value))
			{
				movementOrganization = dataAccess.GetOrganizationRefFromPostalCode(contact.Address1_PostalCode, MemberType.MemberSRK.ToString());
                isMovementMembership = true;
            }

            string breakDateString = dataAccess.GetConfiguration("ME_MembershipYearlyBreakDate");
            string[] breakParts = breakDateString.Split('/');

            DateTime breakDate = new DateTime(DateTime.Now.Year, int.Parse(breakParts[1]), int.Parse(breakParts[0]));

            int regardingYear = DateTime.Now.Year;
            bool startDateNextYear = false;

            if (payment.PaymentDate.Year == breakDate.Year && payment.PaymentDate > breakDate)
            {
                if (dataAccess.ContactHasMemberShipForYear(contact.Id, regardingYear))
                {
                    startDateNextYear = true;
                }
                regardingYear++;
            }

            // organization
            if (string.IsNullOrWhiteSpace(contact.Address1_PostalCode))
                throw new CrmRedyException("No postal code given for payment in CRMMainService.CreateMembershipFromPayment" + Environment.NewLine + ObjectToString.ToString(payment), true);

            srk_organization organization = dataAccess.GetOrganizationFromPostalCode(contact.Address1_PostalCode, membershipTypeEnum.ToString());

            if (organization == null)
                throw new CrmRedyException("No organization could be found for specified contact and postalcode" + Environment.NewLine + ObjectToString.ToString(payment), true);

            // fundraising / campaign code
            if (string.IsNullOrWhiteSpace(payment.CampaignId))
                throw new CrmRedyException("No campaign id given for payment in CRMMainService.CreateMembershipFromPayment" + Environment.NewLine + ObjectToString.ToString(payment), true);

            srk_fundraising campaign = dataAccess.GetFundraisingCampaignFromCampaignId(payment.CampaignId);
            if (campaign == null)
                throw new CrmRedyException("No campaign found for given ID" + Environment.NewLine + ObjectToString.ToString(payment), true);

            // check for future memberships
            int ME_MaxNumberOfFutureMemberships = int.Parse(dataAccess.GetConfiguration("ME_MaxNumberOfFutureMemberships"));
            int futureMemberShips = 0;

            while (dataAccess.ContactHasMemberShipForYear(contact.Id, regardingYear))
            {
                if (futureMemberShips < ME_MaxNumberOfFutureMemberships)
                {
                    regardingYear++;
                    futureMemberShips++;
                }
                else
                {
                    CreateMaxNumberOfFutureMembershipsActivity(payment, contact);
                    return null;
                }
            }

            if (yearIsClosed || futureMemberShips > 0)
            {
                startDateNextYear = true;
            }

            srk_membership newMembership = CreateMembership
                (
					// REF2016CHANTE 26
//					new CrmEntityReference(contact.LogicalName, contact.Id),                            //Contact
					new EntityReference(contact.LogicalName, contact.Id),                            //Contact
					startDateNextYear ? new DateTime(regardingYear, 1, 1) : payment.PaymentDate,        //StartDate
                    new DateTime(regardingYear, 12, 31),                                                //StopDate
                    regardingYear,                                                                      //RegardingYear
					// REF2016CHANTE 26
//					new CrmEntityReference(membershipType.LogicalName, membershipType.Id),              //MembershipType
					new EntityReference(membershipType.LogicalName, membershipType.Id),              //MembershipType
					//REF2016CHANTE 14			Vi låter det smälla vid NULL - precis som förut
//					(decimal)membershipType.srk_price,                                                  //Price
					(decimal)(membershipType.srk_price?.Value),		                                    //Price
					//REF2016CHANTE 14			Vi låter det smälla vid NULL - precis som förut
//					membershipType.srk_price,                                                           //Membership Fee               
					membershipType.srk_price?.Value,													//Membership Fee               
					payment.Amount,                                                                     //Payed amolunt
					// REF2016CHANTE 26
//					new CrmEntityReference(organization.LogicalName, organization.Id),                  //Organization
					new EntityReference(organization.LogicalName, organization.Id),                  //Organization
					movementOrganization,                                                               //OrganizationMembership
					// REF2016CHANTE 26
//					new CrmEntityReference(campaign.LogicalName, campaign.Id),                          //Campaign
					new EntityReference(campaign.LogicalName, campaign.Id),                          //Campaign
                    false,                                                                              //IsPaid
                    payment.PaymentDate,                                                                //PaymendDate
                    (int)PaymentType.Avi,                                                               //PaymentType
                    payment.VoucherNumber,                                                              //VoucherNumber   
                    payment.TransactionNumber,                                                          //TransactionNumber
                    null,                                                                               //MainMember   
                    "",                                                                                 //RecipientName
                    "",                                                                                 //RecipientId
                    "",                                                                                 //RecipientAddressLine1       
                    "",                                                                                 //RecipientAddressLine2
                    "",                                                                                 //RecipientPostalCode   
                    "",                                                                                 //RecipientCity       
                    "",                                                                                 //RecipeintMembershipIncludes                                                                  
                    isMovementMembership,                                                               //IsOrganizationMember
                    null,			// Ingen autogirokoppling kan förekomma..                           //Autogiro   
                    "",                                                                                 //OrderNumber
                    true,                                                                               //AviPrint
                    dataAccess                                                                          //DataAccess
                );

            dataAccess.CreateNote("Medlemskap", "Medlemskap skapat från betalning.", newMembership);
            return newMembership;
        }

        private void CreateMaxNumberOfFutureMembershipsActivity(Payment payment, Contact contact)
        {
            SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                   "Max antal framtida medlemskap uppnått",
                   "En medlemskapsbetalning inkom på en individ som redan har max antal framtida medlemskap. Betalningen måste hanteras manuellt." +
                   Environment.NewLine + "Individ-ID: " + payment.CustomerId +
                   Environment.NewLine + "Belopp: " + payment.Amount +
                   Environment.NewLine + "Kampanjkod: " + payment.CampaignId +
                   Environment.NewLine + "Betaldatum: " + payment.PaymentDate +
                   Environment.NewLine + "OCR: " + payment.PaymentNumber +
                   Environment.NewLine + "Verifikationsnummer AX: " + payment.VoucherNumber +
                   Environment.NewLine + "Transaktionsnummer AX: " + payment.TransactionNumber
                   , contact);
        }


        /// <summary>
        /// Tar emot statuskoder för misslyckade autogirobetalningar BGC spec: 6.2, misslyckats med att dra betalningen. Täckning saknas, inte felkoder om misslyckad upplägg
        /// </summary>
        /// <param name="bizAutogiroPaymetFailed">Objekt för misslyckade autogirobetalningar</param>
        public void ImportAutogiroPaymentFailed(AutogiroResponse bizAutogiroPaymetFailed)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportFailedPayment(bizAutogiroPaymetFailed);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="failMessage"></param>
        public void ImportAutogiroPaymentSetupError(AutogiroPaymentSetupErrorMessage failMessage)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportPaymentError(failMessage);
        }

        /// <summary>
        /// Metod för att ta emot ansökning om autogiroblankett via Redcross (IF0??)
        /// </summary>
        /// <param name="bizAutogiroApplication">Object som innehåller parametrar för autogiroansökan via webben</param>
        public void ImportAutogiroApplicationFromWeb(DataContracts.AutogiroApplication bizAutogiroApplication)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportWebApplication(bizAutogiroApplication);
        }

        /// <summary>
        /// Metod för att ta emot autogiro upplagda via internetbanken (IF028)
        /// </summary>
        /// <param name="bizAutogiroEMandate">Object som innehåller parametrar för autogiro upplagt via internetbank</param>
        public void ImportAutogiroEMandate(DataContracts.AutogiroEMandate bizAutogiroEMandate)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportEMandate(bizAutogiroEMandate);
        }

        /// <summary>
        /// Metod för att ta emot svar på ändringar/makuleringar av betalningsuppdrag från BGC (IF029) BGC spec: 6.5
        /// </summary>
        /// <param name="bizPaymentSpec"></param>
        ///
        public void ImportAutogiroCancellationChangePaymentInCRM(DataContracts.AutogiroResponse bizPaymentSpec)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportPaymentChangeorCancellation(bizPaymentSpec);
        }


        public Contact FindAndUpsertContact(bool lookupMobileNo, string personNo, string firstName, string lastName,
            string streetAddressLine1, string streetAddressLine2, string postalCode, string city, string mobilePhone,
            string phoneNo, string emailAddress, string swishName)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                return dal.FindAndUpsertContact(lookupMobileNo, personNo, firstName, lastName,
                    streetAddressLine1, streetAddressLine2, postalCode, city, mobilePhone,
                    phoneNo, emailAddress, swishName);
            }
        }

		public Guid? FindAndUpsertContactGuid(bool lookupMobileNo, string personNo, string firstName, string lastName,
			string streetAddressLine1, string streetAddressLine2, string postalCode, string city, string mobilePhone,
			string phoneNo, string emailAddress, string swishName)
		{
			return FindAndUpsertContact(lookupMobileNo, personNo, firstName, lastName,
			streetAddressLine1, streetAddressLine2, postalCode, city, mobilePhone,
			phoneNo, emailAddress, swishName)?.Id;
		}


		public Account InsertAccount(string companyNo, string companyName, string firstName, string lastName, string streetAddressLine1, string postalCode, string phoneNo, string emailAddress, string website)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                return dal.InsertAccount(companyNo, companyName, firstName, lastName, streetAddressLine1, postalCode, phoneNo, emailAddress, website);
            }
        }

		public Guid InsertAccountGuid(string companyNo, string companyName, string firstName, string lastName, string streetAddressLine1, string postalCode, string phoneNo, string emailAddress, string website)
		{
			return InsertAccount(companyNo, companyName, firstName, lastName, streetAddressLine1, postalCode, phoneNo, emailAddress, website).Id;
		} 




		/// <summary>
		/// This method should be used when a member need to be created in CRM.
		/// For example when a member is registered on redcross.se(IF003).
		/// </summary>
		/// <param name="membership">Membership-objekt med information om medlemskapet.</param>
		public void ImportMemberInCRM(Membership membership)
        {
            using (ICrmDataAccess dataAccess = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera Membership.
                    if (membership == null)
                        throw new ArgumentNullException("membership");

                    // Validera att medlemstypen är korrekt.
                    if (membership.MemberType != "Member" && membership.MemberType != "YouthMember")
                        throw new ArgumentException("Felaktig medlemstyp specificerad. Giltiga värden: Member,YouthMember", "membership.MemberType");

                    // Validera MemberInfo.
                    if (membership.MemberInfo == null)
                        throw new ArgumentNullException("membership.MemberInfo");

                    // Validera PaymentType.
                    if (membership.PaymentType < 1 || membership.PaymentType > 6)
                        throw new ArgumentException("Felaktig PaymentType, giltiga värden mellan: 1 och 6.", "membership.PaymentType");

                    // OBS! 2013-06-27: Alla medlemskap som först kommer in och skapas upp i CRM ska alltid markeras som obetalda
                    // oavsett om de redan har betalats av kunden. Detta gäller framförallt medlemskap från webben.
                    // Betald-flaggan i CRM sätts till Ja i mottagen betalning efter att medlemskapet har bokförts i Ax och transaktions-
                    // numret har skickats tillbaka.
                    membership.Paid = false;

                    // Mappa till korrekt membertype beroende på om det är ungdomsmedlemskap eller vanligt.
                    if (membership.MemberType == "Member")
                        membership.MemberType = MemberType.MemberSRK.ToString();
                    else if (membership.MemberType == "YouthMember")
                        membership.MemberType = MemberType.MemberRKUF.ToString();

                    // Hitta eller lägg upp ny kontakt.
                    Contact member = dataAccess.FindAndUpsertContact(false, membership.MemberInfo.PersonalNo, membership.MemberInfo.FirstName, membership.MemberInfo.LastName, membership.MemberInfo.PrimaryStreetAddressLine1, membership.MemberInfo.PrimaryStreetAddressLine2, membership.MemberInfo.PrimaryPostalCode, membership.MemberInfo.PrimaryCity, membership.MemberInfo.MobilePhone, membership.MemberInfo.PhoneNo, membership.MemberInfo.EmailAddress, null);

                    if (membership.PaymentType != (int)PaymentType.Avi)
                    {
                        if (dataAccess.ContactHasUnpaidMemberShipCurrentOrFutureYear(member.ContactId.Value))
                            throw new CrmRedyException("Individen har redan ett obetalt medlemskap innevarande eller framtida år.");
                    }

                    //Skapa inte medlemskap för avlidna individer
                    if (dataAccess.ContactIsDeceased(member.ContactId.Value))
                        throw new CrmRedyException("Ett medlemskap kan inte skapas då individen har avslutskod \"Avliden\".");

                    //Skapa inte medlemskap för individer som redan har huvud eller familjemedlemskap innevarande år
                    if (dataAccess.ContactHasMainOrFamilyMembershipCurrentYear(member.ContactId.Value))
                        throw new CrmRedyException("Ett medlemskap kan inte skapas då individen har ett huvud- eller familjemedlemskap innevarande år.");


                    srk_membership srkExistingMembership = null;

                    if (membership.PaymentType == (int)PaymentType.Avi)
                    {
						// REF2016CHANTE 26
//						CrmEntityReference organization = GetOrganization(membership, dataAccess);
						EntityReference organization = GetOrganization(membership, dataAccess);
						//Kontrollera om det redan finns ett medlemskap för innevarande år och organisation.
						srkExistingMembership = dataAccess.CheckExistingMembershipWithContactOrganizationAndYear(member.ContactId.Value, organization.Id, Convert.ToString(membership.OrderDate.Year));
                    }

                    //Om det inte är betalningssätt AVI eller om inget medlemskap för innevarande år på vald organisation hittas, skapa ett nytt.
                    if (srkExistingMembership == null)
                    {
                        // Skapa medlemskap
                        srk_membership srkMemShip = CreateMembershipFromRedCross(membership, dataAccess, member);
                    }
                    //Om medlemskapet redan finns för innevarande år och inkommet medlemskap är av typen AVI.
                    else
                    {
                        HandleExistedAviMemberships(dataAccess, membership, srkExistingMembership);
                    }



                    // OBS: Nedan kod som skapar Avi är bortkommenterat eftersom denna logik nu ligger i pluginen PostMembershipCreate
                    // där avi skapas automatiskt för alla nya medlemskap(FÖ36).

                    //// Skapa Avi-objekt/Betalningsunderlag för alla inkommande förbetalda medlemskap så att synken till Ax-körs i normalflödet.
                    //if (membership.PaymentType != (int)PaymentType.Avi)
                    //{
                    //	dataAccess.CreateAvi(srkMemShip, AviType.NewAviMember);
                    //}
                }
                catch (Exception ex)
                {
                    // TODO: Catch and handle error for each entity create. No transaction support exists for the web service so decide what to do and add manual rollback logic if needed...
                    exception = ex;
                    throw;
                }
                finally
                {
                    try
                    {
                        SrkCrmContext.LogOperation(dataAccess, "CrmMainService::ImportMemberInCRM",
                        ObjectToString.ToString(membership) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }

        private void HandleExistedAviMemberships(ICrmDataAccess dataAccess, Membership importedMembership, srk_membership existedMembership)
        {
            if (existedMembership.srk_ispayed == true)
            {
                dataAccess.CreateNote("Ny avibeställning från web inkom " + importedMembership.OrderDate, "Ett betalt medlemskap finns redan för innevarande år. Inget nytt medlemskap har därför skapats upp.", existedMembership);
            }
            //Om medlemskapets som hittades inte är betalt ändra flaggan på medlemskapet så att ny avisering sker. Skapar även upp en anteckning för detta på medlemskapet.
            else if (existedMembership.srk_ispayed == false)
            {
                existedMembership.srk__avi_print = false;
                dataAccess.CreateNote("Ny avi har beställts", "En ny avi har beställts från webben", existedMembership);
                dataAccess.Update(existedMembership);
            }
        }

        /// <summary>
        /// Registrerar betalning för medlemskap i CRM för angivet OCR-nummer.
        /// </summary>
        /// <param name="payment"></param>
        /// <param name="dataAccess"></param>
        private void UpdateMembershipPaymentInCRM(Payment payment, ICrmDataAccess dataAccess, bool bAutogiroPayment)
        {
            srk_membership membership = null;
            var watchUpdateMembershipPaymentInCRM = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Trace.WriteLine("CRMTRC: Start UpdateMembershipPaymentInCRM: " + watchUpdateMembershipPaymentInCRM.ElapsedMilliseconds);
            bool yearIsClosed = dataAccess.IsYearClosed();

            //Försök hämta ett medlemskap med OCR
            if (!String.IsNullOrEmpty(payment.PaymentNumber))
            {
                membership = dataAccess.GetUnPayedMembershipFromOCRNumber(payment.PaymentNumber);
            }

            //Försök hämta ett medlemskap med hjälp av individid och kampanjkod
            if (membership == null && !String.IsNullOrEmpty(payment.CustomerId) && !String.IsNullOrEmpty(payment.CampaignId))
            {
                membership = dataAccess.GetUnPaidMembershipWithoutOCRNumber(payment.CustomerId, payment.CampaignId, payment.PaymentDate, payment.Amount);
            }

            //Försök hitta betalt årsaviserat huvudmedlemskap som kan ha obetalda familjemedlemmar
            if (membership == null)
            {
                membership = dataAccess.GetPayedMainMembership(payment.CustomerId);
            }

            //Är året stängt så ska vi inte betala på medlemskap för det året
            if (yearIsClosed && membership != null && membership.srk_regardingyear == DateTime.Now.ToLocalTime().Year.ToString())
            {
                membership = null;
            }

            //Kontrollera flaggan för inkorrekt belopp och skapa aktivitet om den är satt till ja
            if (membership != null && membership.srk_payed_amount_incorrect == true)
            {
                CreateIncorrectAmountActivity(membership, payment);
                return;
            }

            //Skapa medlemskap om inget hittas eller om året är stängt
            if (membership == null)
            {
                //Om individen har ett huvud- eller familjemedlemskap innevarande år så antar vi att betalningen gäller ett sådant och skapar en aktivitet och avslutar
                if (dataAccess.ContactHasMainOrFamilyMembershipCurrentYear(payment.CustomerId))
                {
                    Contact contact = dataAccess.GetContact(payment.CustomerId);
                    CreateFamilyActivity(contact, payment);
                    return;
                }
                //Om individen är avliden ska inget nytt medlemskap skapas. En aktivitet skapas och flödet avslutas.
                if (dataAccess.ContactIsDeceased(payment.CustomerId))
                {
                    Contact contact = dataAccess.GetContact(payment.CustomerId);
                    CreateDeceasedMemberActivity(contact, payment);
                    return;
                }

                membership = CreateMembershipFromPayment(payment, dataAccess, yearIsClosed);

                if (membership == null)
                {
                    return;
                }
            }

            #region notneeded
            /*
                // Inga obetalda medlemskap för angivet ocrnr hittades.
             
                Contact contact = null;

                // försök hitta något medlemskap överhuvudtaget till individ med det ocrnumret.
                membership = dataAccess.GetMembershipFromOCRNumber(payment.PaymentNumber);
                
                // om vi fortfarande inte har medlemskap försök hitta på individid
                if (membership == null)
                {


                    // sök fram person på individ id.
                    if (payment.CustomerId.Trim().Length > 0)
                        contact = dataAccess.GetContactFromCustomerId(payment.CustomerId);

                    if (contact != null)
                        membership = dataAccess.GetLatestMembershipForContact(contact);
                }

              
                // om vi hittar en individ men inget medlemskap ska denna person bli medlem.
                // om vi inte hittar en individ gör vi inget mer.
                if (membership == null && contact != null)
                {
                    // om personen inte är medlem, ska de nu bli medlem.

                    Membership memShip = new Membership();
                    memShip.Paid  = true;
                    memShip.PaymentType = (int)PaymentType.Payment;
                    memShip.MemberType = MemberType.Member.ToString();
                    memShip.LocalOrganization = false;
                    memShip.OrderDate = payment.PaymentDate;
                    memShip.PaymentDate = payment.PaymentDate;

                    membership = CreateMembership(memShip, dataAccess, contact);

                    // skapa notering så användaren vet vad som hänt.
                    dataAccess.CreateNote("Medlemskap skapat", "Individen har gjort en inbetalning men hade inget medlemskap som kunde kopplas mot betalningen därför har ett nytt medlemskap skapats.", membership);
                    new FacadeCommon().LogOperation("CrmMainService::UpdateMembershipPaymentInCRM, contact found, no previous membership found, creating a new membership.", null, payment);

                    dataAccess.CreateNote(
                        "Medlemskap betalat.",
                        "Inbetald summa: " + payment.Amount + Environment.NewLine +
                        "Logdata: " + Environment.NewLine +
                        ObjectToString.ToString(payment),
                        membership
                    );

                    return membership;
                }
                else if (contact == null && membership == null)
                {
                    new FacadeCommon().LogOperation("CrmMainService::UpdateMembershipPaymentInCRM, could not find [Contact], no membership created.", null, payment);
                }

            
                // om befintligt medlemskap redan är betalt ska individen få ett nytt medlemskap för nästkommande år.
                if (membership.srk_ispayed == true)
                {
                    new FacadeCommon().LogOperation("CrmMainService::UpdateMembershipPaymentInCRM, membership already paid for.", null, payment);

                    #region notneeded
                    /*
                    Contact contact = dataAccess.GetContact(membership.srk_contactId);
                    if (contact != null)
                    {
                        // lista ut vad nästa år är, eller antag året efter dagens datum
                        int EndYear = DateTime.Now.Year + 1;

                        var latestMembership = dataAccess.GetLatestMembershipForContact(contact);

                        if (latestMembership != null)
                        {
                            if (latestMembership.srk_stopdate.HasValue)
                                EndYear = membership.srk_stopdate.Value.Year + 1;
                        }

                        if (EndYear < DateTime.Now.Year)
                            EndYear = DateTime.Now.Year;

                        Membership memShip = new Membership();
                        memShip.Paid = true;
						memShip.PaymentType = (int)PaymentType.Payment;
                        memShip.MemberType = MemberType.Member.ToString();
                        memShip.LocalOrganization = false;
                        memShip.OrderDate = new DateTime(EndYear, 1, 1); // betalning sätts till första Januari nästa år så att det blir giltigt för hela det året.

                        srk_membership membershipForNextyear = CreateMembership(memShip, dataAccess, contact);

                        // skapa noteringar så användaren vet vad som hänt.
                        dataAccess.CreateNote("Medlemskap skapat", "Det befintliga medlemskapet var redan betalat, därför har ett nytt medlemskap skapats för nästkommande period. Detta är det nya medlemskapet för nästa period. Det gamla medlemskapet har GUID: " + membership.Id.ToString(), membershipForNextyear);
                        dataAccess.CreateNote("Medlemskap skapat", "Det kom in en betalning mot detta medlemskap, men detta medlemskapet var redan betalat. Därför har ett nytt medlemskap skapats för nästkommande period. Det nya medlemskapet har GUID: " + membershipForNextyear.Id.ToString(), membership);

                        dataAccess.CreateNote(
                            "Medlemskap betalat.",
                            "Inbetald summa: " + payment.Amount + Environment.NewLine +
                            "Logdata: " + Environment.NewLine +
                            ObjectToString.ToString(payment),
                            membershipForNextyear
                        );
                    }
                    else
                    {
                        new FacadeCommon().LogOperation("CrmMainService::UpdateMembershipPaymentInCRM, could not find contact", null, membership);
                    }
                     * */
                    #endregion notneeded



            /*
             * Tre typer av betalningar kan komma:
             * 1. Total betalning av ett helt familjemedlemskap.
             * 2. Betalning av ETT medlemskap(vanliga medlemskap + tillkommande familjemedlemskap).
             * 3. Betalning av ETT uppdelat huvudmedlemskap(dvs enbart betalning för huvudmedlem), nytt iom AP117.
             * 4. Betalning på ett betalt OCR som ligger på en årsaviserad huvudmedlem där inkommande belopp motsvarar summan 
             *      av obetalda familjemedlemmar. Då ska logiken slå till för dessa obetalda familjemedlemmar.
             *      Tillägg till betalningslogik för splittrade familjer ala FÖ42.
            */

            if ((membership.srk_ispayed ?? false) == true
                && membership.srk_membershiptypeId != null
                && dataAccess.isMainMemberType(membership.srk_membershiptypeId.Id)
                && membership.srk_startdate.HasValue
                && membership.srk_startdate.Value.ToLocalTime().Month == 1
                && membership.srk_startdate.Value.ToLocalTime().Day == 1)
            {
                // Obetalt, årsaviserat huvudmedlemskap. Kontrollera familjemedlemmarna
                List<srk_membership> family = dataAccess.GetUnpayedFamilyMemberships(membership, Guid.Empty).ToList<srk_membership>();
                decimal familySum = new decimal(0);

                foreach (srk_membership fm in family)
					//REF2016CHANTE 14B
//					familySum += fm.srk_amount.HasValue ? fm.srk_amount.Value : 0;
                    familySum += fm.srk_amount?.Value ?? 0;

                if (familySum == payment.Amount)
                {
                    foreach (srk_membership fm in family)
                    {
                        Payment paymentClone = new Payment()
                        {
							//REF2016CHANTE 14
//							Amount = fm.srk_amount.HasValue ? fm.srk_amount.Value : 0,
							Amount = fm.srk_amount?.Value ?? 0,
							CampaignId = payment.CampaignId,
                            CustomerId = payment.CustomerId,
                            OwnerType = payment.OwnerType,
                            PaymentDate = payment.PaymentDate,
                            PaymentNumber = payment.PaymentNumber,
                            PaymentType = payment.PaymentType,
                            TransactionNumber = payment.TransactionNumber,
                            VoucherNumber = payment.VoucherNumber
                        };
                        UpdateMembershipPayment(paymentClone, fm, dataAccess);
                        SrkCrmContext.Instance.GetSubscriptions(dataAccess).HandleNewMembershipPaymentForContact(fm);
                    }
                }
                else
                {
                    CreateIncorrectFamilyAmountActivity(payment, membership, dataAccess);
                }
            }
            //Befintlig logik som bara letar efter obetalda medlemskap
            else if ((membership.srk_ispayed ?? false) == false)
            {
                UpdateMembershipPayment(payment, membership, dataAccess);

                if (bAutogiroPayment)
                    UpdateAutogiroWithValidPayment(payment.PaymentNumber, payment.PaymentDate);

                SrkCrmContext.Instance.GetSubscriptions(dataAccess).HandleNewMembershipPaymentForContact(membership);

                if ((membership.srk_split_membership ?? false) == false && membership.srk_membershiptypeId != null && dataAccess.isMainMemberType(membership.srk_membershiptypeId.Id))
                {
                    // Betalning har inkommit på ett huvudmedlemskap som INTE är uppdelat(dvs det är en totalbetalning för familjen).
                    // Uppdatera ingående familjemedlemmar. Prenumerationshanteringen för familjemedlemmar hanteras senare per medlem.
                    UpdateFamilymemberForTotalFamilyPayment(payment, membership, dataAccess);
                }
            }
        }


        private void CreateIncorrectFamilyAmountActivity(Payment payment, srk_membership membership, ICrmDataAccess dataAccess)
        {
            SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                 "Inkommande betalning på betalt huvudmedlemskap. Summan av familjemedlemmarna stämmer ej med inkommande betalning.",
                 "En betalning inkom på ett betalt huvudmedlemskap där summan av familjemedlemmarna ej stämmer med inkommande betalning." +
                 Environment.NewLine + "Individ-ID: " + payment.CustomerId +
                 Environment.NewLine + "Belopp: " + payment.Amount +
                 Environment.NewLine + "Kampanjkod: " + payment.CampaignId +
                 Environment.NewLine + "Betaldatum: " + payment.PaymentDate +
                 Environment.NewLine + "OCR: " + payment.PaymentNumber +
                 Environment.NewLine + "Verifikationsnummer AX: " + payment.VoucherNumber +
                 Environment.NewLine + "Transaktionsnummer AX: " + payment.TransactionNumber
                 , membership);
        }

        private void CreateIncorrectAmountActivity(srk_membership membership, Payment payment)
        {
            SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                "Flaggan \"Inbetalt belopp ej korrekt\" är satt.",
                "En betalning inkom på ett medlemskap där flaggan \"Inbetalt belopp ej korrekt\" var satt. Betalningen måste hanteras manuellt." +
                Environment.NewLine + "Individ-ID: " + payment.CustomerId +
                Environment.NewLine + "Belopp: " + payment.Amount +
                Environment.NewLine + "Kampanjkod: " + payment.CampaignId +
                Environment.NewLine + "Betaldatum: " + payment.PaymentDate +
                Environment.NewLine + "OCR: " + payment.PaymentNumber +
                Environment.NewLine + "Verifikationsnummer AX: " + payment.VoucherNumber +
                Environment.NewLine + "Transaktionsnummer AX: " + payment.TransactionNumber
                , membership);
        }

        private void CreateFamilyActivity(Contact contact, Payment payment)
        {
            SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                "En betalning utan OCR inkom på en individ som är huvudmedlem eller familjemedlem innevarande år.",
                "En betalning utan OCR inkom på en individ som är huvudmedlem eller familjemedlem innevarande år. Om betalningen gäller en familj behöver medlemsskapet skapas upp i CRM." +
                Environment.NewLine + "Individ-ID: " + payment.CustomerId +
                Environment.NewLine + "Belopp: " + payment.Amount +
                Environment.NewLine + "Kampanjkod: " + payment.CampaignId +
                Environment.NewLine + "Betaldatum: " + payment.PaymentDate +
                Environment.NewLine + "Verifikationsnummer AX: " + payment.VoucherNumber +
                Environment.NewLine + "Transaktionsnummer AX: " + payment.TransactionNumber
                , contact);
        }

        private void CreateDeceasedMemberActivity(Contact contact, Payment payment)
        {
            SrkCrmContext.Instance.Utility.CreateManualCRMActivity(
                "En betalning som inte matchar något obetalt medlemskap kom in på en avliden individ.",
                "En betalning som inte matchar något obetalt medlemskap kom in på en avliden individ. Betalningen behöver hanteras manuellt." +
                Environment.NewLine + "Individ-ID: " + payment.CustomerId +
                Environment.NewLine + "Belopp: " + payment.Amount +
                Environment.NewLine + "Kampanjkod: " + payment.CampaignId +
                Environment.NewLine + "Betaldatum: " + payment.PaymentDate +
                Environment.NewLine + "Verifikationsnummer AX: " + payment.VoucherNumber +
                Environment.NewLine + "Transaktionsnummer AX: " + payment.TransactionNumber
                , contact);
        }

        /// <summary>
        /// Hanterar medlemskap före betalning registreras
        /// *Inaktiva medlemskap ska aktiveras och avslutskod/avslutsdatum rensas
        /// *Medlemskapet ska förlängas att gälla även nästa år: Om medlemskap gäller innevarande år OCH medlemskapet har ingått i årsavisering (första medlemskap på individen har data) OCH Brytdatum uppnått OCH inget medlemskap finns för nästa år
        /// </summary>
        /// <param name="payment"></param>
        /// <param name="dataAccess"></param>
        private void HandleMembershipBeforePayment(Payment payment, srk_membership membership, ICrmDataAccess dataAccess)
        {
            if (membership.srk_AvslutskodId != null) //Om avslutskod finns
            {
                srk_terminationcodes terminationCode = dataAccess.GetTerminationCodes(membership.srk_AvslutskodId);

                if (terminationCode.srk_resigncode != "A") // Om avslutskod inte är avliden ska medlemskapet aktiveras och avslutskod/avslutsdatum rensas
                // Slå upp avslutskoden
                {
                    membership.srk_AvslutskodId = null;
                    membership.srk_terminationdate = null;

                    // Logga vad som hänt - Medlemskap har återaktiverats
                    dataAccess.CreateNote
                        ("Betalning inkommen på avslutat medlemskap.",
                        "Betalning har inkommit på ett medlemskap som var avslutat, avslutskod och avslutsdatum har nu rensats"
                        + Environment.NewLine + Environment.NewLine + "Logdata:" + Environment.NewLine +
                        ObjectToString.ToString(payment),
                        membership
                    );

                }
                else
                    return; //Om avslutskod är avliden ska medlemskapet inte förlängas så vi avbryter här
            }

			//REF2016CHANTE 14
//			if (membership.srk_regardingyear == DateTime.Now.Year.ToString() && membership.statecode == 1) //Om medlemskapet är avslutat och gäller innevarande år
            if (membership.srk_regardingyear == DateTime.Now.Year.ToString() && (int?)membership.statecode == 1) //Om medlemskapet är avslutat och gäller innevarande år
			{
				membership.srk_stopdate = new DateTime(payment.PaymentDate.Year, 12, 31); //Sätt slutdatum till sista december året för betalning
            }

            if (membership.srk_stopdate != null && membership.srk_stopdate.Value.ToLocalTime().Year == DateTime.Now.Year) //Kontroll: Om medlemskap gäller innevarande år
            {
                if (membership.srk_contactId != null)
                {
                    Contact contact = dataAccess.Get<Contact>(membership.srk_contactId.Id);

                    if (membership.srk_startdate.HasValue
                        && membership.srk_startdate.Value.ToLocalTime().Year == DateTime.Now.Year
                        && (membership.srk_startdate.Value.ToLocalTime().Month == 1
                        && membership.srk_startdate.Value.ToLocalTime().Day == 1) == false) //Kontroll: Om medlemskapet ingått i årsaviseringen ska inte medlemskapet förlängas
                    {

                        membership.srk_startdate = payment.PaymentDate;

                        bool isAfterBreakDate = dataAccess.IsDateAfterYearlyBreakDate(payment.PaymentDate); //Kontroll: Om brytdatum var uppnått när medlemskapet betalades. Jämför brytdatum med betalningsdatum

                        if (isAfterBreakDate)
                        {
                            if (string.IsNullOrWhiteSpace(membership.srk_regardingyear) == true)
                            {
                                string errorMsg = "Fel i: CRMMainService.HandleMembershipBeforePayment, Kunde inte uppdatera medlemskap då avser år saknas ";
                                throw new CrmRedyException(errorMsg, true);
                            }


                            int nextYear = Int32.Parse(membership.srk_regardingyear) + 1;
                            string nextYearString = nextYear.ToString();

                            Guid membershipForNextYear = dataAccess.GetMembershipForYear(membership.Id, nextYearString); //Kontroll: Om det finns medlemsksap för nästa år

                            //Om det inte redan finns medlemskap för nästa år så uppdateras medlemskapet att även gälla nästa år. I annat fall betalas bara befintligt medlemskap
                            if (membershipForNextYear == null || membershipForNextYear == Guid.Empty)
                            {

                                membership.srk_stopdate = new DateTime(membership.srk_stopdate.Value.Year + 1, 12, 31);
                                membership.srk_regardingyear = nextYearString;

                                // Logga vad som hänt - Medlemskapet förlängs att gälla nästa år
                                dataAccess.CreateNote
                                    ("Medlemskap betalat efter brytdatum.",
                                    "Medlemskapet betalades efter brytdatum, därför har medlemskapet förlängts och gäller även nästa år."
                                    + Environment.NewLine + Environment.NewLine + "Logdata:" + Environment.NewLine +
                                    ObjectToString.ToString(payment),
                                    membership
                                );
                            }
                            else
                            {
                                // Logga vad som hänt - Finns redan medlemskap för nästa år
                                dataAccess.CreateNote
                                    ("Medlemskap betalat efter brytdatum.",
                                    "Medlemskapet betalades efter brytdatum men ett medlemskap för nästa år finns redan för individen. Detta medlemskap har därför inte förlängts."
                                    + Environment.NewLine + Environment.NewLine + "Logdata:" + Environment.NewLine +
                                    ObjectToString.ToString(payment),
                                    membership
                                );
                            }


                        }

                    }
                }

            }

            //Om medlemskapet inte gäller innevarande år ska kontroll göras om startdatum ska uppdateras
            if (membership.srk_startdate.HasValue
                         && (membership.srk_startdate.Value.ToLocalTime().Month == 1
                         && membership.srk_startdate.Value.ToLocalTime().Day == 1) == false) //Kontroll: Om medlemskapet ingått i årsaviseringen ska medlemskapet inte förlängas
            {
                membership.srk_startdate = payment.PaymentDate; //Startdatum sätts till betaldatum
            }
        }


        /// <summary>
        /// Registrerar betalning för alla familjemedlemskap i en familj då det sker en
        /// total betalning på huvudmedlemskapet.
        /// </summary>
        /// <param name="payment"></param>
        /// <param name="dataAccess"></param>
        private void UpdateFamilymemberForTotalFamilyPayment(Payment payment, srk_membership mainMembership, ICrmDataAccess dataAccess)
        {
            //Kontrollera att det är ett huvudmedlemskap det handlar om.
            if (mainMembership.srk_membershiptypeId == null)
                return;

            if (!dataAccess.isMainMemberType(mainMembership.srk_membershiptypeId.Id))
                return;

            // Betalt belopp sätts till 0 eftersom hela summan ligger på huvudmedlemmen.
            payment.Amount = 0;

            //Hämta alla kopplade medlemskap.
            List<srk_membership> subMemberships = dataAccess.GetUnpayedMembershipsFromMainMember(mainMembership.Id);

            //Logga om det inte finns några kopplade familjemedlemmar.
            if (subMemberships == null || subMemberships.Count == 0)
            {
                SrkCrmContext.LogOperation(dataAccess, "CrmMainService::UpdateFamilyMembershipPaymentInCRM, no memberships found.",
                        "Det finns inget aktivt och obetalt familjemedlemskap kopplat till huvudmedlemmen." + Environment.NewLine + Environment.NewLine + ObjectToString.ToString(payment),
                        CrmUtility.SystemLogCategory.Integration,
                        CrmUtility.SystemLogType.Warning
                    );
                return;
            }

            //Iterera igenom alla medlemmar i familjen.            
            foreach (srk_membership tmpMember in subMemberships)
            {
                UpdateMembershipPayment(payment, tmpMember, dataAccess, mainMembership);

                SrkCrmContext.Instance.GetSubscriptions(dataAccess).HandleNewMembershipPaymentForContact(tmpMember);
            }
        }

        /// <summary>
        /// Hanterar själva uppdateringen av ett betalt medlemskap.
        /// "Betalt belopp"(srk_payed_amount) uppdateras alltid med vad som kommer från Ax oavsett om det är korrekt eller ej.
        ///		Underliggande familjemedlemskap sätts till 0.
        /// Jämför "Belopp betalningsunderlag" mot inkommet belopp från Ax.
        /// 	□ Vid diff:
        ///			® Sätt flagga "Inbetalt belopp ej korrekt" till true.
        /// 		® Sätt flagga "Betalt" till false.
        /// 		® Låt "Underlag fördelning" vara orörd(null).
        /// 	□ Korrekt belopp:
        /// 		® Sätt flagga "Inbetalt belopp ej korrekt" till false.
        /// 		® Sätt flagga "Betalt" till true.
        /// 		® Sätt "Underlag fördelning" till Medlemskapsavgiften.
        /// </summary>
        private void UpdateMembershipPayment(Payment payment, srk_membership membership, ICrmDataAccess dataAccess, srk_membership mainMembership = null)
        {

            //Logik för att kontrollera om avslutskod/avslutsdatum ska rensas samt om efter brytdatum och medlemskapet ska förlängas att gälla även nästa år
            HandleMembershipBeforePayment(payment, membership, dataAccess);

            //Om individen har ett inaktivt betalt medlemskap innevarande år ska detta aktiveras och avslutskod rensas
            if (membership.srk_regardingyear != DateTime.Now.ToLocalTime().Year.ToString() && !dataAccess.IsDateAfterYearlyBreakDate(DateTime.Now.ToLocalTime()))
                dataAccess.HandleInactivePayedMembershipsCurrentYear(membership.srk_contactId.Id);


            // Kontrollera först om vi har en betalningsdiff. Belopp betalningsunderlag jämförs mot inkommet belopp
            // från Ax. Om det handlar om ett familjemedlemskap så kontrolleras betalning mot huvudmedlemen eftersom
            // det i detta fall gäller en totalbetalning för familjen.
            bool paymentDiff = false;
			//REF2016CHANTE 14
//			decimal membershipAmount = (mainMembership != null ? (mainMembership.srk_amount ?? 0) : (membership.srk_amount ?? 0));
			decimal membershipAmount = (mainMembership != null ? (mainMembership.srk_amount?.Value ?? 0) : (membership.srk_amount?.Value ?? 0));
			//REF2016CHANTE 14
//			decimal paymentAmount = (mainMembership != null ? (mainMembership.srk_payed_amount ?? 0) : payment.Amount);
			decimal paymentAmount = (mainMembership != null ? (mainMembership.srk_payed_amount?.Value ?? 0) : payment.Amount);
			if (membershipAmount != paymentAmount)
                paymentDiff = true;

            if (paymentDiff)
            {
                // Inkommen betalning från Ax diffar mot vad CRM förväntar.
                // Sätt medlemskapet till obetalt och flagga för felaktig betalning.
                membership.srk_ispayed = false;
                membership.srk_payed_amount_incorrect = true;
                membership.srk_distributed_amount = null;
            }
            else
            {
                // Inkommen betalning från Ax är korrekt.
                // Sätt medlemskapet till betalt och släck flagga för felaktig betalning.
                // Sätt även "Underlag fördelning"(srk_distributed_amount) till medlemskapsavgiften.
                membership.srk_ispayed = true;
                membership.srk_payed_amount_incorrect = false;
                membership.srk_distributed_amount = membership.srk_membership_fee;
            }

            membership.srk_paymentdate = payment.PaymentDate;
            membership.srk_voucherno = payment.VoucherNumber;
            membership.srk_transactionno = payment.TransactionNumber;
			//REF2016CHANTE 14
//			membership.srk_payed_amount = payment.Amount;
			membership.srk_payed_amount = new Microsoft.Xrm.Sdk.Money(payment.Amount);


			#region notneeded
			/*
                membership.srk_stopdate = new DateTime(DateTime.Now.Year, 12, 31);

              
                // om betalningen sker efter 1a oktober ska de få hela nästa år, och gratis resten av innevarande år.
                if (payment.PaymentDate.Month >= 10 && membership.srk_contactId != null)
                {
                    // Gör så att nuvarande medlemskap gäller detta året ut
                    membership.srk_startdate = payment.PaymentDate; 
                    membership.srk_stopdate = new DateTime(DateTime.Now.Year, 12, 31);
             
                    // Skapa nytt medlemskap för nästa år.
                    Membership memberDetailsForNextYear = new Membership();
                    memberDetailsForNextYear.Paid = false;
                    memberDetailsForNextYear.PaymentType = (int) PaymentType.Payment; 
                    memberDetailsForNextYear.MemberType = GetNewMembershipTypeFromPreviousMembership(membership);
                    memberDetailsForNextYear.LocalOrganization = GetLocalOrNotFromPreviousMembership(membership);
                    memberDetailsForNextYear.OrderDate = new DateTime(DateTime.Now.Year + 1, 1, 1); // betalning sätts till första Januari nästa år så att det blir giltigt för hela det året.

                    Contact contact = dataAccess.Get<Contact>(membership.srk_contactId.Id);
                    srk_membership membershipForNextyear = CreateMembership(memberDetailsForNextYear, dataAccess, contact);

                    //Sätter det nya medlemskapet till betalt, med samma vouchernummer och betalningsdatum som föregående år.
                    membershipForNextyear.srk_ispayed = true;
                    membershipForNextyear.srk_paymentdate = payment.PaymentDate;
                    membershipForNextyear.srk_voucherno = payment.VoucherNumber;
                    dataAccess.Update(membershipForNextyear);

                    // skapa noteringar så användaren vet vad som hänt.
                    dataAccess.CreateNote
                        ("Medlemskap betalat efter 1a Oktober.",
                        "Medlemskap betalades efter 1a Oktober därför har denna betalning medfört medlemskap även under nästa år. Detta är medlemskapet för nästkommande år. Inbetald summa: " + payment.Amount + Environment.NewLine + Environment.NewLine +
                        "Logdata: " + Environment.NewLine +
                        ObjectToString.ToString(payment),
                        membershipForNextyear
                    );

                    dataAccess.CreateNote
                        ("Medlemskap betalat efter 1a Oktober.",
                        "Medlemskap betalades efter 1a Oktober därför har denna betalning medfört medlemskap även under nästa år. Detta är medlemskapet för resterande del av året. Inbetald summa: " + payment.Amount + Environment.NewLine + Environment.NewLine +
                        "Logdata: " + Environment.NewLine +
                        ObjectToString.ToString(payment),
                        membership
                    );
                }
                */
			#endregion notneeded

			dataAccess.Update(membership);

            if (membership.statecode != 0 && membership.srk_terminationdate == null) //Aktivera medlemskapet om det har status inaktivt och avslutsdatum har rensats
            {
                dataAccess.SetEntityStateCode(membership, 0, 1);
            }



            if (paymentDiff)
            {
                // Logga diff på betalningssumman
                string details = "";
                if (mainMembership != null)
                    details = "Familjemedlemskap betalat med DIFF, se mer info på anteckning för huvudmedlemmen.";
                else
					//REF2016CHANTE 14
//					details = "Medlemskap betalat med DIFF, inbetald summa: " + payment.Amount + ", belopp betalningsunderlag: " + (membership.srk_amount ?? 0);
					details = "Medlemskap betalat med DIFF, inbetald summa: " + payment.Amount + ", belopp betalningsunderlag: " + (membership.srk_amount?.Value ?? 0);
				dataAccess.CreateNote
                    ("Medlemskap betalat.",
                    details + Environment.NewLine + Environment.NewLine +
                    "Logdata: " + Environment.NewLine +
                    ObjectToString.ToString(payment),
                    membership
                );
            }
            else
            {
                // Logga vad som betalats
                dataAccess.CreateNote
                    ("Medlemskap betalat.",
                    "Medlemskap betalat, inbetald summa: " + payment.Amount + Environment.NewLine + Environment.NewLine +
                    "Logdata: " + Environment.NewLine +
                    ObjectToString.ToString(payment),
                    membership
                );

                UpdateContactDataAfterMembershipPayment(membership, payment, dataAccess);
            }
        }


        private void UpdateContactDataAfterMembershipPayment(srk_membership membership, Payment payment, ICrmDataAccess dataAccess)
        {
            // update first membership day on contact.
            DateTime paymentDate = payment.PaymentDate;
            if (membership.srk_contactId != null)
            {
                // ska inte längra sättas, görs via Workflows
                //Contact contact = dataAccess.Get<Contact>(membership.srk_contactId.Id);
                //if (contact.srk_date_first_membership > paymentDate || contact.srk_date_first_membership == null)
                //{
                //    contact.srk_date_first_membership = paymentDate;
                //    dataAccess.Update(contact);
                //}
            }
        }

        private bool GetLocalOrNotFromPreviousMembership(srk_membership membership)
        {
            bool localOrganization = false;

            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                if (membership.OrganizationId != null)
                {
                    var organization = dal.GetOrganization(membership.OrganizationId);
                    if (organization.srk_organizationtype.Value == (int)OrganizationType.CentralOrganization)
                        localOrganization = true;
                }
            }

            return localOrganization;
        }

        /// <summary>
        /// Metod för att ta emot svar på medgivanden från BGC (IF028) BGC spec: 6.3
        /// </summary>
        /// <param name="bizMandate"></param>
        public void ImportAutogiroMandateInCRM(DataContracts.AutogiroResponse bizMandate)
        {
            var autogiro = new CRMAutogiro();
            autogiro.ImportMandateResponse(bizMandate);
        }

        /// <summary>
        /// Lägger upp en order i CRM.
        /// Används bland annat i integrationsflöde IF024 - produktköp på redcross.se.
        /// </summary>
        /// <param name="order">Order-objekt med information och minst en tillhörande orderrad.</param>
        public void ImportOrderInCRM(DataContracts.Order order)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // NOTE: Denna integration kör enbart mot WebForms i nuläget och behöver sedan ändras när den nya eHandeln är på plats.
                    // Från webforms kan enbart privatpersoner beställa engångsköp av juletiketter. Prenumerationer hanteras genom
                    // manuellt upplägg i CRM. DataContracts-klasserna har tagit höjd för den nya eHandeln med stöd för orderkorgstänk etc
                    // men det kan bara vara en orderrad som anger antal juletiketter.

                    /*
                        1. Validera parametrar.
                        2. Hitta Individ, konto eller krets som gjort köpet och koppla denna till ordern. Om ingen träff finns
                           så ska individen/konto/krets läggas upp. OBS endast privatperson för WebForms.
                        3. Skapa huvudorder med kopplad kampanj och en orderrad med juletiketts-produkten.
                        4. Skapa gåva med samma kampanj och koppla ihop med produktköpet.
                    */

                    // Validera Order.
                    if (order == null)
                        throw new ArgumentNullException("order");

                    // I webForms ska det bara vara en orderrad.
                    if (order.OrderLines == null || order.OrderLines.Count != 1)
                        throw new ArgumentException("Ordern måste ha EN orderrad.", "order.OrderLines");

                    // OrderOwnerType styr vem som gjort ordern. 
                    // P=Private person, F=Company, O=Organization
                    if (order.OrderOwnerType != "P"/*Endast privatperson giltig med WebForms*/ )
                        throw new ArgumentException("Felaktig OrderOwnerType, giltigt värde: P.", "order.OrderOwnerType");

                    // Validera PaymentType.
                    if (order.PaymentType < 1 || order.PaymentType > 6)
                        throw new ArgumentException("Felaktig PaymentType, giltiga värden mellan: 1 och 6.", "order.PaymentType");

                    // Validera BuyerInfo.
                    if (order.BuyerInfo == null)
                        throw new ArgumentNullException("order.BuyerInfo");

                    Contact contact = null;
                    switch (order.OrderOwnerType)
                    {
                        case "P":
                            {
                                contact = dal.FindAndUpsertContact(false, order.BuyerInfo.PersonalNo, order.BuyerInfo.FirstName, order.BuyerInfo.LastName, order.BuyerInfo.PrimaryStreetAddressLine1, order.BuyerInfo.PrimaryStreetAddressLine2, order.BuyerInfo.PrimaryPostalCode, order.BuyerInfo.PrimaryCity, order.BuyerInfo.MobilePhone, order.BuyerInfo.PhoneNo, order.BuyerInfo.EmailAddress, null);
                                break;
                            }
                        case "F":
                            {
                                // Företag. Inte aktuellt från WebForms men kommer att bli med nya eHandeln(EPI).
                                break;
                            }
                        case "O":
                            {
                                // Organisationer. Inte aktuellt från WebForms men kommer att bli med nya eHandeln(EPI).
                                break;
                            }
                    }

                    // Hämta produktid för julklappsetiketter från konfigurationen.                    
                    string configKey = "JKE_ProductID_ChristmasStickers";
                    string strProductNo = dal.GetConfiguration(configKey);

                    srk_product product = dal.FindProduct(strProductNo);
                    if (product == null)
                        throw new Exception(string.Format("Produkten för Julklappsetiketter, med id {0}, kunde inte hittas.", strProductNo));

                    // Hämta kampanjkod
                    srk_fundraising fundraising = dal.GetFundraisingCampaignFromCampaignId(order.CampaignId);
					// REF2016CHANTE 26
//					CrmEntityReference fundraisingRef = null;
					EntityReference fundraisingRef = null;
					if (fundraising != null)
						// REF2016CHANTE 26
//						fundraisingRef = new CrmEntityReference(fundraising.LogicalName, fundraising.Id);
                        fundraisingRef = new EntityReference(fundraising.LogicalName, fundraising.Id);

                    // Lägg upp produktköp
                    var purchase = new srk_productpurchase()
                    {
                        srk_ordernumber = order.OrderNo,
                        srk_orderdate = order.OrderDate,
                        srk_paymentdate = order.PaymentDate,
						//REF2016CHANTE 14
//						srk_paymenttype = order.PaymentType,
						srk_paymenttype = new Microsoft.Xrm.Sdk.OptionSetValue(order.PaymentType),
						srk_deliverydate = order.EstimatedReceiveDate,
                        srk_deliverysrk = order.DeliveryDate,
						//REF2016CHANTE 14
//						srk_amount = order.TotalAmount,
						srk_amount = new Microsoft.Xrm.Sdk.Money(order.TotalAmount),
						srk_ispaid = order.Paid,
						//REF2016CHANTE 14
//						srk_sourcesystem = (int)OrderSourceSystem.Web,
						srk_sourcesystem = new Microsoft.Xrm.Sdk.OptionSetValue((int)OrderSourceSystem.Web),
						//REF2016CHANTE 14
//						srk_customertype = (int)CustomerType.Contact,
						srk_customertype = new Microsoft.Xrm.Sdk.OptionSetValue((int)CustomerType.Contact),
						srk_fundraising_id = fundraisingRef
                    };

					// Koppla kontakt till produktköp
					// REF2016CHANTE 26
//					purchase.srk_contactId = new CrmEntityReference(Contact.EntityLogicalName, contact.Id);
					purchase.srk_contactId = new EntityReference(Contact.EntityLogicalName, contact.Id);

					// Spara produktköpet
					dal.Create(purchase);

                    // Skapa produkt i produktköp
                    // Det finns bara en rad från webforms för julklappsetiketter och det är bara Antal vi hämtar.
                    OrderLine line = order.OrderLines[0];
                    int quantity = (int)line.Quantity;
                    srk_ProductInOrder purchaseLine = new srk_ProductInOrder()
                    {
                        srk_name = product.srk_name,
                        srk_quantity = quantity,
						//REF2016CHANTE 14
//						srk_price = (product.srk_price.HasValue ? quantity * product.srk_price : 0),
						srk_price = new Microsoft.Xrm.Sdk.Money((product.srk_price != null ? quantity * product.srk_price.Value : 0)),
						srk_itemnumber = product.srk_axitemnumber,
						// REF2016CHANTE 26
//						srk_productId = new CrmEntityReference(product.LogicalName, product.Id),
//                        srk_ProductsInOrder = new CrmEntityReference(purchase.LogicalName, purchase.Id)
                        srk_productId = new EntityReference(product.LogicalName, product.Id),
                        srk_ProductsInOrder = new EntityReference(purchase.LogicalName, purchase.Id)
                    };

                    if (product.TransactionCurrencyId != null)
                    {
						// REF2016CHANTE 26
//						purchaseLine.TransactionCurrencyId = new CrmEntityReference(TransactionCurrency.EntityLogicalName, product.TransactionCurrencyId.Id);
						purchaseLine.TransactionCurrencyId = new EntityReference(TransactionCurrency.EntityLogicalName, product.TransactionCurrencyId.Id);
					}

					// Spara orderraden
					dal.Create(purchaseLine);

                    // Skapa Betalningsunderlag/avi-objekt för alla inkommande förbetalda julklappsetiketter så att synken till Ax-körs i normalflödet.
                    if (order.PaymentType != (int)PaymentType.Avi)
                    {
                        // NOTE: Bortkommenterat 2013-06-28: Inkommande produktköp är idag bara julklappsetiketter
                        // och dessa ska läggas upp som produktköp SAMT en gåva. Underlaget för produktköpet ska däremot
                        // inte skickas till Ax vilket styrs av flaggan srk_invoicestatus som nu inte sätts längre nedan.

                        // Uppdatera flagga för att skicka iväg fakturaunderlag för ordern till BizTalk.
                        // purchase.srk_invoicestatus = true;
                        // dataAccess.Update(purchase);

                        // Skapa och fyll gåvo-objekt
                        Donation donationObject = new Donation()
                        {
                            DonationOwnerType = order.OrderOwnerType,
                            DonationDate = order.OrderDate,
                            PaymentDate = null/*Betaldatum sätts till null och uppdateras efter bokning i Ax.*/,
                            Paid = false/*Gåvor sätts alltid som obetalda och bockas av efter bokföring/betalning i Ax*/,
                            DonationType = (int)DonationType.CommonGift,
                            PaymentType = order.PaymentType,
                            TotalAmount = order.TotalAmount,
                            CampaignId = order.CampaignId,
                            DonorInfo = order.BuyerInfo,
                            OrderNumber = order.OrderNo
                        };

                        // Skapa gåva
                        srk_donation gift = CreateDonation(donationObject, null, dal, contact, purchase, true);
                        gift.srk_name = SrkCrmContext.Instance.Utility.GetDefaultNameForEntity(dal, gift);

                        //Om klumpbokning är aktiverat ska inget betalningsunderlag skapas för julklappsetiketter
                        if (!IsClumpBookDonation(dal, donationObject))
                        {
                            // Skapa Betalningsunderlag/avi-objekt
                            dal.CreateAvi(gift);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // TODO: Catch and handle error for each entity create.
                    // No transaction support exists for the web service so decide what to do and add manual rollback logic if needed...
                    exception = ex;
                    throw;
                }
                finally
                {
                    try
                    {
                        SrkCrmContext.LogOperation(dal, "CrmMainService::ImportOrderInCRM",
                        ObjectToString.ToString(order) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }

		/// <summary>
		/// Lägger upp en gåva i CRM.
		/// Används bland annat i integrationsflöde IF022 - registrera gåva på redcross.se
		/// och IF051 - ta emot SMS-betalning från LinkMobility - telefongåva från Viatel.
		/// </summary>
		/// <param name="donation">Donation-objekt med information om gåvan.</param>
		public void ImportDonationInCRM(Donation donation)
        {
            using (ICrmDataAccess dataAccess = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera Donation.
                    if (donation == null)
                        throw new ArgumentNullException("donation");

                    // DonationOwnerType styr typ av givare.
                    // P=Privatperson, F=Företag, A=Anonym privatperson
                    if (donation.DonationOwnerType != "P" && donation.DonationOwnerType != "A" && donation.DonationOwnerType != "F")
                        throw new ArgumentException("Felaktig DonationOwnerType, giltigt värde: P, A eller F.", "donation.DonationOwnerType");

                    // Är det ett SMS med "STOPP" i Eventkod eller Prefix så ska denna "donation" specialbehandlas, då det inte är en donation
                    if ((donation.PaymentType == (int)PaymentType.SMS || donation.PaymentType == (int)PaymentType.SMSmanadsdragning) &&
                        (donation.CampaignPrefix.Contains("STOP") || donation.CampaignEventCode.Contains("STOP")))
                    {
                        Contact donor = dataAccess.GetContactByMobile(donation.DonorInfo.MobilePhone);

                        if (donor != testMerge) { 
                            dataAccess.HandleSMS_STOPP(donor, donation.DonationDate);
                        }
                        else { 
                            SrkCrmContext.LogOperation(dataAccess, "Kunde inte hitta individ", $"Individ med telefonnummer {donation.DonorInfo.MobilePhone} har skickat STOPP, men kunde inte hittas i CRM", CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Error);
                        }
                        return;
                    }

                    // Validera gåvotyp
                    if (! (donation.DonationType >= 1 && donation.DonationType <= 7) )
                        throw new ArgumentException("Felaktig DonationType, giltiga värden mellan 1 och 7.", "donation.DonationType");

                    // Validera betalsätt
                    if (! ((donation.PaymentType >= 1 && donation.PaymentType <= 12) || donation.PaymentType == 16) )
						throw new ArgumentException("Felaktig PaymentType, giltiga värden mellan 1 och 12 samt 16.", "donation.PaymentType");

                    // Validera DonorInfo.
                    if (donation.DonorInfo == null)
                        throw new ArgumentNullException("donation.DonorInfo");

                    //Validera Swish-specifika fält, kontrollera att Swish-numret är upplagt samt kontrollera att Swish-gåva inte hämtats förut
                    if (donation.PaymentType == (int)PaymentType.Swish)
                    {
                        if (String.IsNullOrEmpty(donation.CampaignEventCode))
                            throw new ArgumentException("Fältet Swishnummer är obligatoriskt", "donation.CampaignEventCode");

                        if (String.IsNullOrEmpty(donation.DonationReferenceNo))
                            throw new ArgumentException("Fältet transaktionsnummer Swish är obligatoriskt", "donation.DonationReferenceNo");

                        if (dataAccess.GetSwishNumberIdBySwishNumber(donation.CampaignEventCode) == null)
                            throw new ArgumentException("Swishnumret som betalningen avser är inte upplagd i CRM.");

                        if (dataAccess.GetDonationIdByReferenceNumber(donation.DonationReferenceNo, (int)PaymentType.Swish) != null)
                            throw new ArgumentException("Swish-gåvan med angivet transaktionsnummer finns redan upplagd i CRM.");
                    }

                    //Validera att kampanjkod finns och är aktivt. Vid felaktig kampanjkod eller SMS-kod så skickas felet tillbaka till BizTalk.
                    srk_fundraising camp = null;
                    if (donation.PaymentType == (int)PaymentType.SMS || donation.PaymentType == (int)PaymentType.SMSmanadsdragning)
                    {
						camp = dataAccess.GetValidCampaignFromSMSChannel(donation.CampaignPrefix, donation.CampaignEventCode, donation.DonationDate);
						// SMS-gåvor och SMS-månadsdragningar kommer in som PaymentType.SMS båda två. 
						// Vi behöver titta på Kampanjens GåvoTyp får att se om det är fråga om SMSmanadsdragning
						if (camp?.srk_gift_type?.Value == (int?)srk_gift_type_picklist.ManadsgavaSMS)
						{
							donation.PaymentType = (int)PaymentType.SMSmanadsdragning;
						}
                    }
                    else if (donation.PaymentType == (int)PaymentType.Swish)
                    {
                        camp = dataAccess.GetValidCampaignFromSwishNumber(donation.CampaignEventCode, donation.DonationDate);
                    }
                    else
                    {
                        camp = dataAccess.GetValidCampaignFromCampaignId(donation.CampaignId, donation.DonationDate);
                    }


                    if (camp == null)
                        throw new ArgumentException(String.Format
                            ("Det gick inte att hämta en kampanjkod som uppfyller kraven: Kampanjstatus ej lika med Under upprättande och Gåvodatum ({0}) mellan start och slutdatum för kampanjen.", donation.DonationDate.ToString("d")));

                    //Klumpbokning 
                    bool bookAsLump = IsClumpBookDonation(dataAccess, donation);
                    //Klumpbokning

                    // OBS! 2013-06-11: Alla gåvor som först kommer in och skapas upp i CRM ska alltid markeras som obetalda
                    // oavsett om de redan har betalats av kunden. Detta gäller framförallt gåvor från webben och SMS/Telefon.
                    // Betald-flaggan i CRM sätts till Ja i mottagen betalning efter att gåvan har bokförts i Ax och transaktions-
                    // numret har skickats tillbaka.
                    donation.Paid = false;
                    donation.PaymentDate = null;

                    // Skapa gåva
                    srk_donation gift = CreateDonation(donation, null, dataAccess, null, null, false, bookAsLump);
                    gift.srk_name = SrkCrmContext.Instance.Utility.GetDefaultNameForEntity(dataAccess, gift);

                    // Skapa Avi-objekt/Betalningsunderlag för alla inkommande förbetalda gåvor så att synken till Ax-körs i normalflödet.
                    if (donation.PaymentType != (int)PaymentType.Avi && donation.TotalAmount > 0 /*Klumpbokning SMS*/ && !(bookAsLump) /*Klumpbokning SMS*/)
                    {
                        dataAccess.CreateAvi(gift);
                    }
                }
                catch (Exception ex)
                {
                    exception = ex;
                    throw;
                }
                finally
                {
                    try
                    {
                        if (donation.DonationOwnerType == "A")
                        {
                            if (donation.DonorInfo == null)
                                donation.DonorInfo = new Person();

                            donation.DonorInfo.FirstName = "Okänd";
                            donation.DonorInfo.LastName = "Givare";
                        }
                        SrkCrmContext.LogOperation(dataAccess, "CrmMainService::ImportDonationInCRM",
                            ObjectToString.ToString(donation) +
                            Environment.NewLine + Environment.NewLine +
                            exception,
                       CrmUtility.SystemLogCategory.Integration,
                       exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }

        //Klumpbokning 
        private bool IsClumpBookDonation(ICrmDataAccess dataAccess, Donation donation)
        {
            bool result = false;
            if (donation.PaymentType != (int)PaymentType.Payment && donation.PaymentType != (int)PaymentType.Autogiro && donation.PaymentType != (int)PaymentType.Avi)
            {
                return Convert.ToBoolean(int.Parse(dataAccess.GetConfiguration("IsClumpActivated")));
            }
            return result;
        }
        //Klumpbokning 

        /// <summary>
        /// Importerar utskrifter som skall göras lokalt i CRM
        /// </summary>
        /// <param name="postalDelivery">Utskicksobjekt</param>
        public void ImportPostalDeliveryInCRM(PostalDelivery postalDelivery)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera
                    if (postalDelivery == null)
                        throw new ArgumentNullException("postalDelivery");

                    //log if print to print service
                    if (postalDelivery.PrintType == PrintType.PrintService)
                    {
                        dal.CreateLetterActivityOnCustomer(postalDelivery.CustomerId, postalDelivery.DispatchType, postalDelivery.EventDate, postalDelivery.CampaignCode, postalDelivery.DeliveryLogId);
                        Guid info30Guid;

                        //check deliverycode and info30 is guid of productpurchase
                        if (Guid.TryParse(postalDelivery.Info30, out info30Guid)) //hantera produktköpsaktivitet
                        {
                            //dataAccess.CreateLetterActivityOnProductPurchase(info30Guid, postalDelivery.DispatchType, DateTime.Now, postalDelivery.CampaignCode);


                            //Om utskickstypen som kommer in innebär att det är JKE som  har levererats ska produktköpet uppdateras till levererat
                            //utskickstyper för leverans specas i konfig: JKE_DispatchLogTypeForDeliveryOfStickers

                            string validDispatchTypes = dal.GetConfiguration("JKE_DispatchLogTypeForDeliveryOfStickers");

                            string[] dispatchTypes = validDispatchTypes.Split(';');

                            var christmasStickersDelivered = false;

                            foreach (string s in dispatchTypes)
                            {
                                if (postalDelivery.DispatchType.Equals(s.Trim(), StringComparison.CurrentCultureIgnoreCase))
                                {
                                    christmasStickersDelivered = true;
                                }
                            }

                            if (christmasStickersDelivered)
                            {
                                dal.SetProductPurchaseDeliveryStatus(info30Guid, true, DateTime.Now);
                            }
                        }
                    }
                    else
                    {

                        //CRM cannot handle mindate: 0001/01/01
                        if (postalDelivery.PaymentDate == DateTime.MinValue)
                            postalDelivery.PaymentDate = new DateTime(1900, 1, 1);

                        dal.CreatePostalDelivery(
                            postalDelivery.EventDate,
                            postalDelivery.DeliveryCode,
                            postalDelivery.DeliveryLogId,
                            postalDelivery.CustomerId,
                            postalDelivery.DispatchType,
                            postalDelivery.DispatchTypeName,
                            postalDelivery.Amount,
                            postalDelivery.PaymentDate,
                            postalDelivery.CampaignCode,
                            postalDelivery.OCR,
                            postalDelivery.Autogiro,
                            postalDelivery.Info1,
                            postalDelivery.Info2,
                            postalDelivery.Info3,
                            postalDelivery.Info4,
                            postalDelivery.Info5,
                            postalDelivery.Info6,
                            postalDelivery.Info7,
                            postalDelivery.Info8,
                            postalDelivery.Info9,
                            postalDelivery.Info10,
                            postalDelivery.Info11,
                            postalDelivery.Info12,
                            postalDelivery.Info13,
                            postalDelivery.Info14,
                            postalDelivery.Info15,
                            postalDelivery.Info16,
                            postalDelivery.Info17,
                            postalDelivery.Info18,
                            postalDelivery.Info19,
                            postalDelivery.Info20,
                            postalDelivery.Info21,
                            postalDelivery.Info22,
                            postalDelivery.Info23,
                            postalDelivery.Info24,
                            postalDelivery.Info25,
                            postalDelivery.Info26,
                            postalDelivery.Info27,
                            postalDelivery.Info28,
                            postalDelivery.Info29,
                            postalDelivery.Info30);
                    }
                }
                catch (CrmDuplicateException dex)
                {
                    exception = dex;
                }
                catch (Exception ex)
                {
                    // NOTE: Loggning sker nu bara på fel i ImportPostalDeliveryInCRM eftersom det kan komma +200' poster
                    // in till denna på en gång från Neolane. Detta snabbar upp flödet något.
                    exception = ex;
                    SrkCrmContext.LogOperation(dal, "CrmMainService::ImportPostalDeliveryInCRM",
                        ObjectToString.ToString(postalDelivery) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        CrmUtility.SystemLogType.Error);
                    throw;
                }
                finally
                {
                    try
                    {                        
                        //new FacadeCommon().LogOperation("CrmMainService::ImportPostalDeliveryInCRM", exception, postalDelivery);
                    }
                    catch { }
                }
            }
        }


        /// <summary>
        /// Importerar emailaktiviteter på kunder
        /// </summary>
        /// <param name="emailLog">logobjekt</param>
        public void ImportDispatchLog(DispatchLog dispatchLog)
        {
            using (ICrmDataAccess dataAccess = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera
                    if (dispatchLog == null)
                        throw new ArgumentNullException("dispatchLog");

                    switch (dispatchLog.Type)
                    {
                        case DataContracts.DispatchType.Email:
                            dataAccess.CreateEmailActivityOnCustomer(dispatchLog.CustomerId, dispatchLog.EventDate, dispatchLog.URLMirrorPage, dispatchLog.CampaignCode, dispatchLog.DeliveryCode);
                            break;
                        case DataContracts.DispatchType.SMS:
                            dataAccess.CreateSMSActivityOnCustomer(dispatchLog.CustomerId, dispatchLog.EventDate, dispatchLog.CampaignCode, dispatchLog.DeliveryCode);
                            break;
                        case DataContracts.DispatchType.TMCall:
                            dataAccess.CreateTMCallActivityOnCustomer(dispatchLog.CustomerId, dispatchLog.EventDate, dispatchLog.CampaignCode, dispatchLog.DeliveryCode);
                            break;
                    }

                }
                catch (CrmDuplicateException ex)
                {
                    // Koll om det är CrmDuplicatException fel och kasta inte i sådan fall vidare utan logga bara (i finally)
                    exception = ex;                    
                }
                catch (Exception ex)
                {
                    exception = ex;
                    // Throw to Biztalk
                    throw;
                }
                finally
                {
                    try
                    {
                        SrkCrmContext.LogOperation(dataAccess, "CrmMainService::ImportDispatchLog",
                        ObjectToString.ToString(dispatchLog) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Importerar emailaktiviteter på kunder
        /// </summary>
        /// <param name="emailLog">logobjekt</param>
        public void ImportOptOut(OptOutInfo optOutInfo)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                try
                {
                    // Validera
                    if (optOutInfo == null)
                        throw new ArgumentNullException("optOutInfo");


                    switch (optOutInfo.Type)
                    {
                        case DataContracts.DispatchType.Email:
                            dal.RegisterOptOut(optOutInfo.CustomerId, optOutInfo.EmailAddress, Common.DispatchType.Email, optOutInfo.OptOutStatus);
                            break;
                        case DataContracts.DispatchType.SMS:
                            dal.RegisterOptOut(optOutInfo.CustomerId, optOutInfo.MobileNumber, Common.DispatchType.SMS, optOutInfo.OptOutStatus);
                            break;
                    }

                }
                catch (Exception ex)
                {
                    exception = ex;
                    throw;
                }
                finally
                {
                    try
                    {
                        SrkCrmContext.LogOperation(dal, "CrmMainService::ImportOptOut",
                        ObjectToString.ToString(optOutInfo) +
                        Environment.NewLine + Environment.NewLine +
                        exception,
                        CrmUtility.SystemLogCategory.Integration,
                        exception != null ? CrmUtility.SystemLogType.Warning : CrmUtility.SystemLogType.Information);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Metod som bearbetar filen med optin och optouts som prenumeranter på Digitala utskick.
        /// Arbetet sker i en egen tråd så att BT ska kunna få tillbaka ett svar, att meddelandet är mottaget, innan arbetet är slutfört.
        /// </summary>
        /// <param name="filePath">Sökväg till den fil som innehåller OptIn/OptOuts</param>
        public void EmailSubsriberOptInOut(string filePath)
        {
            System.ComponentModel.BackgroundWorker BackgroundWorker = new System.ComponentModel.BackgroundWorker();

            BackgroundWorker.DoWork +=
                new System.ComponentModel.DoWorkEventHandler(Background_ImportOptInOptout);

            SRK.Crm.Main.Common.OptInOptOut.ImportOptInOptOut importClass = new Common.OptInOptOut.ImportOptInOptOut();
            importClass.filepath = filePath;

            // Start the asynchronous operation.  
            BackgroundWorker.RunWorkerAsync(importClass);
        }

        private void Background_ImportOptInOptout(
            object sender,
            System.ComponentModel.DoWorkEventArgs e)
        {
            SRK.Crm.Main.Common.OptInOptOut.ImportOptInOptOut importClass = (SRK.Crm.Main.Common.OptInOptOut.ImportOptInOptOut)e.Argument;
            
            importClass.ImportFile();
        }

        /// <summary>
        /// Metod som hanterar fil från SPAR rad för rad 
        /// Parsar fil
        /// Startar systemjobb
        /// </summary>
        /// <param name="filePath"></param>
        public void ImportSPARFile(string filePath)
        {
            Guid systemjobid;
            
            using (ICrmDataAccess dal = SrkCrmContext.Instance.CreateDataAccess())
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    SrkCrmContext.LogOperation(dal, GetType().Name, string.Format("File path is invalid ('{0}'", filePath), CrmUtility.SystemLogCategory.Common, CrmUtility.SystemLogType.Error);
                    throw new ArgumentException("value cannot be null or whitespace", "filePath");
                }

                if (!File.Exists(filePath))
                {
                    SrkCrmContext.LogOperation(dal, GetType().Name, string.Format("File '{0}' does not exist, or user does not have enough privileges", filePath), CrmUtility.SystemLogCategory.Common, CrmUtility.SystemLogType.Error);
                    throw new ArgumentException("File does not exist, or user does not have enough privileges", filePath);
                }

                using (Stream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (stream.Length == 0)
                    {
                        SrkCrmContext.LogOperation(dal, GetType().Name, string.Format("File '{0}' has zero length", filePath), CrmUtility.SystemLogCategory.Common, CrmUtility.SystemLogType.Error);
                        throw new ArgumentException("File has zero length", "filePath");
                    }
                    string fileName = Path.GetFileName(filePath);

                    SystemJobInformation jobInfo = new SystemJobInformation
                    {
                        JobType = Common.SystemJobType.SPARImport,
                        ExtraInfo = new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>("Filsökväg", filePath) }
                    };

                    systemjobid = SrkCrmContext.Instance.SystemJob.InitializeSystemJob(dal, jobInfo);
                    SrkCrmContext.LogOperation(dal, GetType().Name, string.Format("SPAR system job started for file '{0}'", Path.GetFileName(filePath)), CrmUtility.SystemLogCategory.Common, CrmUtility.SystemLogType.Information);
                    // Systemjobb skapat och detta i sin tur triggar igång ett CustomWorkflow (SPARImport), som plockar filnamnet och bearbetar filen
                    
                }
            }
        }

        /// <summary>
        /// Uppdaterar eller skapar en verksamhet
        /// </summary>
        /// <param name="operation"></param>
        public void UpsertOperation(Operation operation)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {
                    // BUGG 2371 - Sök reda på befintlig Verksamhet baserat på ID-nummer istället för FrivilligGUID om inkommande FrivilligGUID har färre än 7 tecken.
                    List<srk_operation> result;
                    if(operation.FrivilligId.Length > 6)
                        result = dal.Context.srk_operationSet.Where(x => x.srk_frivilligguid == operation.FrivilligId).ToList<srk_operation>();
                    else
                        result = dal.Context.srk_operationSet.Where(x => x.srk_idnumber == operation.FrivilligId).ToList<srk_operation>();

                    int count = result.Count;
                    if (count == 0)
                    {
                        // Verksamhet fanns inte sedan innan, så den skapas i CRM                        
                        srk_operation newOperation = new srk_operation();
                        if (operation.ClosureDate.HasValue) {
                            newOperation.srk_terminationdate = operation.ClosureDate;
                            // Avslutskod hårdkodad till "Verksamhet upphört" om den inte kommer med från Frivillig
                            EntityReference termCode = null;
                            try
                            {
                                termCode = operation.TerminationCodeId != null ? new EntityReference(srk_terminationcodes.EntityLogicalName, operation.TerminationCodeId) : dal.GetTerminationCode("VERK_TERM").ToEntityReference();
                            }
                            catch (Exception e)
                            {
                                SrkCrmContext.LogOperation("Misslyckades att hämta avslutskod för verksamhet:" + operation.Name, e);
                            }
                            newOperation.srk_terminationId = termCode;
                        }
                        newOperation.srk_emailaddress_ope = operation.Email;
                        newOperation.srk_name = operation.Name;
                        newOperation.srk_description = operation.Description;
                        // Obligatoriskt med id-nummer så det sätts ett dummyvärde innan räknare funkar.
                        // newOperation.srk_idnumber = "f" + DateTime.Now.Ticks.ToString().Substring(DateTime.Now.Ticks.ToString().Length-5);

                        if(operation.MainOrganization != null && operation.MainOrganization.Id != null) {
                            // Letar upp mainorg via ID som kommer in från Frivillig
                            var orgs = dal.Context.srk_organizationSet.Where(o => o.srk_idnumber == operation.MainOrganization.Id).ToList<srk_organization>();
                            if(orgs.Count == 1) {
                                newOperation.srk_legalorganization = orgs.First().ToEntityReference();
                            }
                        }
                        newOperation.srk_frivilligguid = operation.FrivilligId;
                        if(operation.OperationType != null && operation.OperationType.Guid != null)
                            newOperation.srk_operationtypeId = new EntityReference(srk_operation_type.EntityLogicalName, operation.OperationType.Guid);
                        else { 
                            // Har det inte kommit in ngn typ så sätts den till FVK Testverksamhetstyp
                            newOperation.srk_operationtypeId = new EntityReference(srk_operation_type.EntityLogicalName, new Guid("FBC690FE-81C4-E711-80DE-00505681402E"));
                        }
                        dal.CreateOperation(newOperation);
                    }
                    if(count == 1)
                    {
                        // Verksamheten hittades och uppdateras
                        srk_operation changedOperation = result.First();
                        if (operation.ClosureDate != null)
                        {
                            changedOperation.srk_terminationdate = operation.ClosureDate;
                            // Avslutskod hårdkodad till "Verksamhet upphört" om den inte kommer med från Frivillig
                            EntityReference termCode = operation.TerminationCodeId != null ? new EntityReference(srk_terminationcodes.EntityLogicalName, operation.TerminationCodeId) : dal.GetTerminationCode("VERK_TERM").ToEntityReference();
                            changedOperation.srk_terminationId = termCode;

                            // Avslutar alla uppdrag som tillhör denna verksamhet
                            try
                            {
                                dal.EndAllFrivilligAssignments(changedOperation);
                            }
                            catch (Exception e)
                            {
                                SrkCrmContext.LogOperation("Misslyckades att avsluta uppdrag för verksamhet:" + changedOperation.srk_idnumber, e);
                            }
                        }
                        changedOperation.srk_name = operation.Name;
                        changedOperation.srk_description = operation.Description;
                        changedOperation.srk_emailaddress_ope = operation.Email;

                        if (operation.MainOrganization != null && operation.MainOrganization.Id != null)
                        {
                            // Letar upp mainorg via ID som kommer in från Frivillig
                            var orgs = dal.Context.srk_organizationSet.Where(o => o.srk_idnumber == operation.MainOrganization.Id).ToList<srk_organization>();
                            if (orgs.Count == 1)
                            {
                                changedOperation.srk_legalorganization = orgs.First().ToEntityReference();
                            }
                        }
                        changedOperation.srk_frivilligguid = operation.FrivilligId;
                        if (operation.OperationType != null && operation.OperationType.Guid != null)
                            changedOperation.srk_operationtypeId = new EntityReference(srk_operation_type.EntityLogicalName, operation.OperationType.Guid);
                        dal.UpdateOperation(changedOperation);
                    }
                    else if(count > 1)
                    {
                        // Fler än en Verksamhet hittad --> FEL
                        throw new Exception(String.Format("Angivet ID-nummer är inte unikt. {0} verksamheter med samma ID-nummer: {1}.", count, operation.Id.ToString()));
                    }
                    
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CrmMainService::UpsertOperation", e, operation);
                }
            }
        }

        /// <summary>
        /// Uppdaterar eller skapar ett uppdrag, anropas ifrån Frivillig (via BizTalk)
        /// </summary>
        /// <param name="operation"></param>
        public void UpsertAssignment(Assignment assignment)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {
                    // Hämtar verksamhet för uppdraget
                    List<srk_operation> operations = dal.Context.srk_operationSet.Where(x => x.srk_frivilligguid == assignment.OperationFrivilligId).ToList<srk_operation>();
                    srk_operation connectedOperation;
                    if (operations.Count == 1)
                        connectedOperation = operations.First();
                    else
                    {
                        throw new Exception(String.Format("Angivet ID-nummer för verksamhet, från Frivillig, kan inte hittas: {0}",  assignment.OperationFrivilligId));
                    }
                    // Hämtar Individ för uppdraget
                    List<Contact> contacts = dal.Context.ContactSet.Where(c => c.srk_frivilligguid == assignment.Contact.FrivilligGuid).ToList<Contact>();
                    if(contacts.Count != 1)
                    {
                        throw new Exception(String.Format("Angivet ID-nummer för Individ, från Frivillig, kan inte hittas: {0} ", assignment.Contact.FrivilligGuid));
                    }
                    Contact contactWhoHasAssignment = contacts.First();

                    // Tar reda på uppdragstyp utifrån assignment.AssignmentType.Name, fallback till FRIVILLIG
                    srk_assignmenttype assignmentType;
                    try
                    {
                        assignmentType = assignment.AssignmentType.Name != null && !assignment.AssignmentType.Name.Equals("") ? dal.GetAssignmentType(assignment.AssignmentType.Name) : dal.GetAssignmentType("FRIVILLIG");
                    }
                    catch(Exception)
                    {
                        assignmentType = dal.GetAssignmentType("FRIVILLIG");
                    }

                    // Letar upp samma uppdrag baserat på contact, verksamhet och uppdragstyp
                    List<srk_assignment> assignments = dal.Context.srk_assignmentSet.Where(x => x.srk_contactassignmentId.Id == contactWhoHasAssignment.ContactId 
                                        && x.srk_operation_id.Id == connectedOperation.Id
                                        && x.statecode == (int)StateCode.Active
                                        && x.srk_assignmenttypeId.Id == assignmentType.Id).ToList<srk_assignment>();

                    int count = assignments.Count;
                    if (count == 0)
                    {
                        // Nytt uppdrag som inte funnits tidigare i CRM
                        srk_assignment newAssignment = new srk_assignment();
                        string campaignid = "10490"; //Default
                        try
                        {
                            campaignid = dal.GetConfiguration("Frivillig_CampaignIdAssignment");
                        }
                        catch(Exception)
                        {
                            SrkCrmContext.LogOperation(dal, "Konfig värde saknas: Frivillig_CampaignIdAssignment", "Använder default kampanj 10490 för uppdraget som skapas", CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Warning);
                        }
                        newAssignment.srk_fundraising_id = dal.GetFundraisingCampaignFromCampaignId(campaignid).ToEntityReference();

                        newAssignment.srk_assignmenttypeId = assignmentType.ToEntityReference();

                        newAssignment.srk_startdate = assignment.StartDate;
                        newAssignment.srk_enddate = assignment.TerminationDate;
                        newAssignment.srk_contactassignmentId = contactWhoHasAssignment.ToEntityReference();

                        newAssignment.srk_OrganizationId = connectedOperation.srk_legalorganization; // new EntityReference(srk_organization.EntityLogicalName, assignment.Organization.Guid);
                        newAssignment.srk_operation_id = connectedOperation.ToEntityReference();// new EntityReference(srk_operation.EntityLogicalName, assignment.OperationFrivilligId);
                        newAssignment.statecode = srk_assignmentState.Active;
                        dal.CreateAssignment(newAssignment);
                    }
                    if (count == 1)
                    {
                        // Uppdraget hittat i CRM, så det uppdateras
                        srk_assignment changedAssignment = assignments.First();
                        
                        changedAssignment.srk_startdate = assignment.StartDate;
                        changedAssignment.srk_enddate = assignment.TerminationDate;                       
                        changedAssignment.srk_OrganizationId = connectedOperation.srk_legalorganization; // new EntityReference(srk_organization.EntityLogicalName, assignment.Organization.Guid);
                        changedAssignment.srk_operation_id = connectedOperation.ToEntityReference();// new EntityReference(srk_operation.EntityLogicalName, assignment.OperationFrivilligId);
                        
                        // Bortkommenterat då uppdrag inte kan avslutas från Frivillig, utan det görs i CRM
                        // då datumet enddate passerats eller om verksamhet/organisation läggs ned
                        //if (assignment.TerminationDate != null && assignment.TerminationDate > DateTime.Now)
                        //{
                        //    // Avslutar uppdrag
                        //    string termcodestring = "AVSLFRIVIL"; // default
                        //    try
                        //    {
                        //        termcodestring = dal.GetConfiguration("Frivillig_TerminationCode_Uppdrag");
                        //    }
                        //    catch (Exception)
                        //    {
                        //        SrkCrmContext.LogOperation(dal, "Konfig värde saknas: Frivillig_TerminationCode_Uppdrag", "Använder default 'AVSLFRIVIL'", CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Warning);
                        //    }

                        //    srk_terminationcodes termCode = dal.GetTerminationCode(termcodestring);
                        //    changedAssignment.srk_terminationcodes_id = termCode.ToEntityReference();
                        //    changedAssignment.srk_terminationdate = assignment.TerminationDate;
                        //    changedAssignment.srk_terminationreason = assignment.TerminationReason;
                        //    changedAssignment.srk_enddate = assignment.TerminationDate;
                        //    changedAssignment.statecode = srk_assignmentState.Inactive;
                        //}
                        dal.UpdateAssignment(changedAssignment);
                    }
                    else if (count > 1)
                    {
                        // Fler än ett Uppdrag hittat --> FEL
                        throw new Exception(String.Format("Angivet ID-nummer är inte unikt. {0} uppdrag med samma ID-nummer (från Frivillig): {1}.", count, assignment.FrivilligGuid));
                    }
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CrmMainService::UpsertAssignment", e, assignment);
                }
            }
        }

        /// <summary>
        /// Uppdaterar eller skapar en Contact baserat på information ifrån Frivillig
        /// Är det en ny individ som inte finns i CRm plockas alla uppgifter men om den redan finns så sätts bara
        /// Email och Mobilnummer om.
        /// </summary>
        /// <param name="personSimple"></param>
        public void UpsertPersonSimple(PersonSimple personSimple)
        {   
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {
                    //Använder Contact-entitet som parameter och behållare för värden
                    Contact c = new Contact();
                    c.srk_EmailRedcross = personSimple.Email;
                    c.FirstName = personSimple.FirstName;
                    c.srk_frivilligguid = personSimple.Identifier;
                    //c.srk_status = personSimple.IsValid;
                    c.LastName = personSimple.LastName;
                    c.srk_socialsecuritynumber = personSimple.PersonNumber;
                    c.MobilePhone = personSimple.Phone;
                    c.srk_date_first_relation = personSimple.Created; // Bara en behållare av createddate från Frivillig
                    //c.srk_status = personSimple.Status;
                    
                    Contact newContact = dal.GetFrivilligContact(c);
                    // Sätter email och telefonnummer från Frivillig ifall individen fortfarande är aktiv, annars är de nollställda
                    if (personSimple.Status == true) { 
                        newContact.srk_EmailRedcross = personSimple.Email;
                        newContact.MobilePhone = personSimple.Phone;
                    }
                    newContact.srk_frivilligguid = personSimple.Identifier;

                    // Om det är en NY individ (prospekt samt skapad i CRM nyss (CreatedOn == null), så att det inte är ett gammal prospekt) som skapats så ska address_source sättas
                    if (newContact.srk_status.Value == (int)ContactStatus.Prospect && newContact.CreatedOn == null)
                    {
                        newContact.srk_addres_source = "Skapad i frivillig";
                    }
                    // om relstatus = inaktiv och avslutskod är Avliden så ska inget göras i CRM
                    if (newContact.srk_status.Value == (int)ContactStatus.Inactive && newContact.srk_terminationCodeId.Id == dal.GetTerminationCode("A").Id)
                    {
                        SrkCrmContext.LogOperation(dal,"Aktiverar inte avlidna", "Individ med ID: " + newContact.srk_contactid + " är avliden men försöktes återupplivas av Frivillig. Inte tillåtet.", CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Warning);
                        return;
                    }
                    // om relstatus = inaktiv och aktiv från Frivillig så ska termdate och termcode nollställas samt anteckning skapas
                        if (newContact.srk_status.Value == (int)ContactStatus.Inactive && personSimple.Status == true)
                    {
                        newContact.srk_termination_date = null;
                        newContact.srk_terminationCodeId = null;
                        dal.CreateNote("Individen har aktiverats", "Denne individ har skapat ett konto i Frivillig och därmed aktiverats", newContact);
                    }

                    // Sätter skapad i frivillig till NYTT datum om denna individ inte redan finns i Frivillig
                    if (!newContact.srk_existsinfrivillig.HasValue || newContact.srk_existsinfrivillig.HasValue && !newContact.srk_existsinfrivillig.Value)
                        newContact.srk_createdinfrivillig = personSimple.Created > DateTime.MinValue ? personSimple.Created : DateTime.Now;

                    // Om det kommer in en avaktiverad personSimple så ska den "tas bort" från Frivillig genom att sätta Flaggan.
                    newContact.srk_existsinfrivillig = personSimple.Status;
                    // Alla uppdrag som den individen har ska också avslutas om individen avslutats i Frivillig
                    if (personSimple.Status == false)
                    {
                        try { 
                            dal.EndAllFrivilligAssignments(newContact);
                        }
                        catch(Exception e)
                        {
                            SrkCrmContext.LogOperation("Misslyckades att avsluta uppdrag för individ:" + newContact.srk_contactid, e);
                        } 
                    }
                    // Om personen är aktiv och inte möjligdublett så ska relationsstatus vara aktiv i CRM
                    if (personSimple.Status && newContact.srk_status.Value != (int)ContactStatus.Duplicate)
                        newContact.srk_status = new Microsoft.Xrm.Sdk.OptionSetValue((int)ContactStatus.Active);

                    dal.Update(newContact);
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CrmMainService::UpsertPersonSimple", e, personSimple);
                }
            }
        }

        /// <summary>
        /// Returnerar en lista med PersonSimple av Contact som har avslutats sedan <updatedAfter>
        /// och finns i Frivillig
        /// </summary>
        /// <param name="updatedAfter"></param>
        /// <returns></returns>
        public List<PersonSimple> GetAllTerminatedPersonSimple(DateTime updatedAfter)
        {
            List<PersonSimple> simpleContactsToReturn = new List<PersonSimple>();
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {
                    
                    List<Contact> teminatedContacts = dal.GetFrivilligTerminatedContacts(updatedAfter);
                    PersonSimple pc;
                    foreach (Contact c in teminatedContacts)
                    {
                        pc = new PersonSimple();
                        pc.FirstName = c.FirstName;
                        pc.LastName = c.LastName;
                        pc.Identifier = c.srk_frivilligguid;
                        pc.PersonNumber = c.srk_socialsecuritynumber;
                        pc.Status = false;
                        simpleContactsToReturn.Add(pc);
                    }
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CRMMainService:GetAllTerminatedPersonSimple", e);
                }
                return simpleContactsToReturn;
            }
        }

        /// <summary>
        /// Returnerar en lista med PersonSimple av Contact som har fått permanent personnummer sedan <updatedAfter>
        /// och finns i Frivillig
        /// </summary>
        /// <param name="updatedAfter"></param>
        /// <returns></returns>
        public List<PersonSimple> GetAllPermanentSocSecNumberPersonSimple(DateTime updatedAfter)
        {
            List<PersonSimple> simpleContactsToReturn = new List<PersonSimple>();
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {

                    List<Contact> newSocSecNumberContacts = dal.GetFrivilligNewSecSocNumberContacts(updatedAfter);
                    PersonSimple pc;
                    foreach (Contact c in newSocSecNumberContacts)
                    {
                        pc = new PersonSimple();
                        pc.FirstName = c.FirstName;
                        pc.LastName = c.LastName;
                        pc.Identifier = c.srk_frivilligguid;
                        pc.PersonNumber = c.srk_socialsecuritynumber;
                        pc.Status = c.srk_existsinfrivillig.HasValue && c.srk_existsinfrivillig.Value;
                        simpleContactsToReturn.Add(pc);
                    }
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CRMMainService:GetAllPermanentSocSecNumberPersonSimple", e);
                }
                return simpleContactsToReturn;
            }
        }

        /// <summary>
        /// Uppdaterar en person enligt uppgifter från SPAR eller Hitta.
        /// </summary>
        /// <param name="sparContact">Kontakt från SPAR med uppgifter som ska uppdateras.</param>
        public void UpdateContactFromSPAR(SPARContact sparContact)
        {
            Exception exception = null;
            srk_systemjob job;
            try
            {
                if (sparContact == null)
                    throw new ArgumentNullException("sparContact");

                // Validera att vi har satt SystemJobId
                if (sparContact.SystemJobId == null || sparContact.SystemJobId == Guid.Empty)
                    throw new ArgumentException("SystemJobId måste innehålla en Guid till aktuellt systemjobb i CRM.", "sparContact.SystemJobId");

                // Validera att vi har korrekta radnummer för aktuell uppdatering
                if (sparContact.RowNumber < 1 || sparContact.TotalNumberofRows < 1)
                    throw new ArgumentException("RowNumber och TotalNumberofRows måste vara större än 0.", "sparContact.RowNumber, sparContact.TotalNumberofRows");
                if (sparContact.RowNumber > sparContact.TotalNumberofRows)
                    throw new ArgumentException("RowNumber är större än TotalNumberofRows.", "sparContact.RowNumber");

                // Validera att vi en korrekt nyckel i ContactAddressId
                if (string.IsNullOrWhiteSpace(sparContact.ContactAddressId))
                    throw new ArgumentException("Nyckeln ContactAddressId får ej vara blankt.", "sparContact.ContactAddressId");

                string[] arrKeys = sparContact.ContactAddressId.Split('|');
                if (arrKeys.Length != 2)
                    throw new ArgumentException("Nyckeln ContactAddressId måste bestå av ID för kontakten och adressen på formatet; IndividID|AddressGuid.", "sparContact.ContactAddressId");

                using (ICrmDataAccess dal = new CrmDataAccess())
                {
                    // Validera att systemjobbet är skapat
                    try
                    {
                        job = dal.Get<srk_systemjob>(sparContact.SystemJobId);
                    }
                    catch
                    {
                        throw new ArgumentException("Systemjobbet med angivet SystemJobId finns inte i CRM.", "sparContact.SystemJobId");
                    }

                    try
                    {
                        CRMSPARImport spar = new CRMSPARImport(dal);
						//REF2016CHANTE 14
//						spar.UpdateContactFromSPAR(sparContact, arrKeys, job.srk_type);
						spar.UpdateContactFromSPAR(sparContact, arrKeys, job.srk_type?.Value);
					}
					catch (Exception ex)
                    {
                        // Om felet beror på uppdatering av en rad i SPAR-importen så svälj felet
                        // och logga det i finally nedan.
                        exception = ex;
                    }
                }
            }
            catch (Exception exMain)
            {
                // Om felet beror på felaktiga parametrar etc kasta tillbaka det till klienten.
                exception = exMain;
                throw;
            }
            finally
            {
                // Logga bara fel i SPAR-importen för att öka prestanda.
                if (exception != null)
                    SrkCrmContext.LogOperation("CrmMainService::UpdateContactFromSPAR", exception, sparContact, sparContact.SystemJobId);
            }
        }

        /// <summary>
        /// Uppdaterar epost från utbildningswebben
        /// </summary>
        /// <param name="contactId"></param>
        /// <param name="epostAddress"></param>
        public void UpdateContactFromCourseWeb(Guid contactId, string epostAddress)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                try
                {
					// REF2016CHANTE 26
//					CrmEntityReference crmref = new CrmEntityReference { Id = contactId };
					EntityReference crmref = new EntityReference { Id = contactId };
					var contact = dal.GetContact(crmref);
                    if (contact.EMailAddress1 != epostAddress)
                    {
                        contact.EMailAddress1 = epostAddress;
                        dal.Update(contact);
                    }
                }
                catch (Exception e)
                {
                    SrkCrmContext.LogOperation("CrmMainService::UpdateContactFromCourseWeb", e, contactId);
                }
            }
        }

        /// <summary>
        /// Uppdaterar postnummer enlig uppgifter från Postens postnummerfil.
        /// </summary>
        /// <param name="postalCode"></param>
        public void UpdatePostalCode(PostalCode postalCode)
        {


            try
            {
                if (postalCode == null)
                    throw new ArgumentNullException("postalCode");

                // Validera att vi har satt SystemJobId
                if (postalCode.SystemJobId == Guid.Empty)
                    throw new ArgumentException("SystemJobId måste innehålla en Guid till aktuellt systemjobb i CRM.", "postalCode.SystemJobId");

                // Validera att vi har korrekta radnummer för aktuell uppdatering
                if (postalCode.RowNumber < 1 || postalCode.TotalNumberofRows < 1)
                    throw new ArgumentException("RowNumber och TotalNumberofRows måste vara större än 0.", "postalCode.RowNumber, postalCode.TotalNumberofRows");
                if (postalCode.RowNumber > postalCode.TotalNumberofRows)
                    throw new ArgumentException("RowNumber är större än TotalNumberofRows.", "postalCode.RowNumber");

                //Validera övriga fält
                if (string.IsNullOrWhiteSpace(postalCode.PostalNumber))
                    throw new ArgumentException("Postnummer får ej vara blankt.", "postalCode.PostalNumber");
                if (string.IsNullOrWhiteSpace(postalCode.City))
                    throw new ArgumentException("Orten får ej vara blankt.", "postalCode.City");
                if (string.IsNullOrWhiteSpace(postalCode.MunicipalityCode))
                    throw new ArgumentException("Kommunkoden får ej vara blankt.", "postalCode.MunicipalityCode");
                if (string.IsNullOrWhiteSpace(postalCode.MunicipalityName))
                    throw new ArgumentException("Namnet på kommunen får ej vara blankt.", "postalCode.MunicipalityName");
                if (string.IsNullOrWhiteSpace(postalCode.CountyCode))
                    throw new ArgumentException("Länskoden får ej vara blankt.", "postalCode.CountyCode");
                if (string.IsNullOrWhiteSpace(postalCode.CountyName))
                    throw new ArgumentException("Länet på länet får ej vara blankt.", "postalCode.CountyName");


                ImportPostalCode ipc = new ImportPostalCode();
                ipc.UpsertPostalCode(new PostalCodeInfo()
                {
                    City = postalCode.City,
                    CountyCode = postalCode.CountyCode,
                    CountyName = postalCode.CountyName,
                    MunicipalityCode = postalCode.MunicipalityCode,
                    MunicipalityName = postalCode.MunicipalityName,
                    PostalNumber = postalCode.PostalNumber,
                    RowNumber = postalCode.RowNumber,
                    SystemJobId = postalCode.SystemJobId,
                    TotalNumberofRows = postalCode.TotalNumberofRows
                });

                // Kontroller om det är sista postnummret som importeras.
                if (postalCode.RowNumber == postalCode.TotalNumberofRows)
                {
                    FinalizeSystemJob(new DataContracts.SystemJobInformation()
                    {
                        JobType = SystemJobType.PostalCodeImport,
                        SystemJobId = postalCode.SystemJobId,
                        NumberOfRecords = postalCode.TotalNumberofRows,
                        EndTime = DateTime.Now
                    });
                }
            }
            catch (Exception e)
            {
                SrkCrmContext.LogOperation("CrmMainService::UpdatePostalCode", e, postalCode, postalCode.SystemJobId);
            }
        }

        /// <summary>
        /// Initierar och skapar ett Systemjobb i CRM.
        /// </summary>
        /// <param name="sysInformation"></param>
        public Guid InitializeSystemJob(DataContracts.SystemJobInformation sysInformation)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                Guid? jobId = null;
                try
                {
                    if (sysInformation == null)
                        throw new ArgumentNullException("sysInformation");
                    if (TranslateJobType(sysInformation.JobType) == Common.SystemJobType.Unknown)
                        throw new ArgumentException("sysInformation.JobType");

                    jobId = SrkCrmContext.Instance.SystemJob.InitializeSystemJob(dal, new SystemJobInformation()
                    {
                        SystemJobId = sysInformation.SystemJobId,
                        JobType = TranslateJobType(sysInformation.JobType),
                        StartTime = sysInformation.StartTime,
                        EndTime = sysInformation.EndTime,
                        NumberOfRecords = sysInformation.NumberOfRecords,
                        ExtraInfo = sysInformation.ExtraInfo
                    });

                    return jobId.Value;
                }
                catch (Exception ex)
                {
                    exception = ex;
                    throw;
                }
                finally
                {
                    SrkCrmContext.LogOperation("CrmMainService::InitializeSystemJob", exception, (sysInformation), jobId.Value);
                }
            }
        }

        /// <summary>
        /// Avslutar Systemjob
        /// </summary>
        /// <param name="sysInformation"></param>
        public void FinalizeSystemJob(DataContracts.SystemJobInformation sysInformation)
        {
            using (ICrmDataAccess dal = new CrmDataAccess())
            {
                Exception exception = null;
                Guid? jobId = null;
                try
                {
                    if (sysInformation == null)
                        throw new ArgumentNullException("sysInformation");

                    jobId = SrkCrmContext.Instance.SystemJob.FinalizeSystemJob(dal, new SystemJobInformation
                    {
                        SystemJobId = sysInformation.SystemJobId,
                        JobType = TranslateJobType(sysInformation.JobType),
                        StartTime = sysInformation.StartTime,
                        EndTime = sysInformation.EndTime,
                        NumberOfRecords = sysInformation.NumberOfRecords,
                        ExtraInfo = sysInformation.ExtraInfo
                    });
                }
                catch (Exception ex)
                {
                    exception = ex;
                    throw;
                }
                finally
                {
                    SrkCrmContext.LogOperation("CrmMainService::FinalizeSystemJob", exception, sysInformation, (sysInformation != null && sysInformation.SystemJobId != null) ? sysInformation.SystemJobId.Value : jobId.GetValueOrDefault());
                }
            }
        }

        #region Private helper functions

        /// <summary>
        /// Metod som uppdaterar autogirouppdraget vid inkommen betalning
        /// </summary>
        /// <param name="referenceNumber">Betalningsuppdragets referensnummer</param>
        /// <param name="dataAccess">CrmDataAccess</param>
        private srk_autogiro UpdateAutogiroWithValidPayment(string referenceNumber, DateTime? paymentDate)
        {
            var autogiroHandler = new CRMAutogiro();
            srk_autogiro autogiro = autogiroHandler.HandleNewPayment(referenceNumber, paymentDate);
            return autogiro;
        }

        /// <summary>
        /// Metod som tar in ett bankgironummer och kontrollerar om det är bgnr för gåva eller medlem
        /// </summary>
        /// <param name="creditorBankgiroNr">Bankgironummer</param>
        /// <param name="dataAccess">CrmDataAccess</param>
        /// <returns>Retunerar true om medlem och false om gåva. Kastar fel om bankgironummer inte kan hittas </returns>

        /// <summary>
        /// Skapar ny gåva och sätter korrekt givartyp (P, F, O, A)
        /// </summary>
        /// <param name="donation"></param>
        /// <param name="payment"></param>
        /// <param name="dataAccess"></param>
        /// <returns></returns>
        private srk_donation CreateDonation(Donation donation, Payment payment, ICrmDataAccess dataAccess, Contact useContact = null, srk_productpurchase productPurchase = null, bool externalNotice = false, bool bookAsLump = false)
        {
            string debugNoteText = "";
            string customerId = "", voucherNumber = "", transactionNumber = "", paymentNumber = "";
            Microsoft.Xrm.Sdk.Entity negativeDonor = null;


            if (payment != null)
            {
                customerId = payment.CustomerId;
                voucherNumber = payment.VoucherNumber;
                transactionNumber = payment.TransactionNumber;
                paymentNumber = payment.PaymentNumber;
            }

            debugNoteText = "Gåva skapad: '" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "'" + Environment.NewLine;
            debugNoteText += "DonationType: '" + Enum.GetName(typeof(DonationType), donation.DonationType) + "'" + Environment.NewLine;

            // Skapa gåva och sätt nödvändig info.
            srk_donation gift = new srk_donation()
            {
                srk_donationdate = donation.DonationDate,
				//REF2016CHANTE 14
//				srk_amount = donation.TotalAmount,
				srk_amount = new Microsoft.Xrm.Sdk.Money(donation.TotalAmount),
				srk_ispayed = donation.Paid,
				//REF2016CHANTE 14
//				srk_paymenttype = donation.PaymentType,
				srk_paymenttype = new Microsoft.Xrm.Sdk.OptionSetValue(donation.PaymentType),
				srk_paymentdate = donation.PaymentDate,
                srk_voucherno = voucherNumber,
                srk_transactionno = transactionNumber,
                srk_referencenumber = (donation.DonationReferenceNo != null && donation.DonationReferenceNo.Length > 0 ? donation.DonationReferenceNo : donation.OrderNumber),   //donation.DonationReferenceNo               
                srk_ocrnumber = (donation.Paid ? paymentNumber : ""), // Om gåvan är markerad som betald så sätts paymentNumber annars genererar CRM ett OCRnr i Create-pluginen.
                srk_external_notice = externalNotice,
                srk_ordernumber = donation.OrderNumber,
                //Om det är Swish, SMS-gåva eller SMSmånadsdragning skall gåvodatum och tid sättas.
                srk_gifttimestamp = (donation.PaymentType == (int)PaymentType.Swish 
								|| donation.PaymentType == (int)PaymentType.SMS 
								|| donation.PaymentType == (int)PaymentType.SMSmanadsdragning ? donation.DonationDate : (DateTime?)null)
            };

            // Sätt korrekt givartyp
            if (donation.DonationOwnerType == "P" || donation.DonationOwnerType == "A")
				//REF2016CHANTE 14
//				gift.srk_donor = (int)CustomerType.Contact;
                gift.srk_donor = new Microsoft.Xrm.Sdk.OptionSetValue((int)CustomerType.Contact);
            else if (donation.DonationOwnerType == "F")
				//REF2016CHANTE 14
//				gift.srk_donor = (int)CustomerType.Account;
                gift.srk_donor = new Microsoft.Xrm.Sdk.OptionSetValue((int)CustomerType.Account);
            else if (donation.DonationOwnerType == "O")
				//REF2016CHANTE 14
//				gift.srk_donor = (int)CustomerType.Organization;
                gift.srk_donor = new Microsoft.Xrm.Sdk.OptionSetValue((int)CustomerType.Organization);

            debugNoteText += "DonationOwnerType: '" + donation.DonationOwnerType + "'" + Environment.NewLine;
            debugNoteText += "PaymentType: '" + Enum.GetName(typeof(PaymentType), donation.PaymentType) + "'" + Environment.NewLine;

            if (gift.srk_donor != null)
				// REF2016CHANTE 38
//				debugNoteText += "CustomerType: '" + Enum.GetName(typeof(CustomerType), gift.srk_donor) + "'" + Environment.NewLine;
                debugNoteText += "CustomerType: '" + Enum.GetName(typeof(CustomerType), gift.srk_donor?.Value) + "'" + Environment.NewLine;

            // Koppla till rätt produktköp
            if (productPurchase != null)
            {
				// REF2016CHANTE 26
//				gift.srk_productgiftid = new CrmEntityReference(srk_productpurchase.EntityLogicalName, productPurchase.Id);
				gift.srk_productgiftid = new EntityReference(srk_productpurchase.EntityLogicalName, productPurchase.Id);
			}

			// Hitta eller lägg upp ny privatperson.
			if (donation.DonationOwnerType == "P")
            {
                // Sök upp individ på mobilnr vid SMS/Telefon-gåva.
                Contact donor = null;

                if (useContact != null) /*Sök bara upp kontakt om vi inte redan skickat in en som ska användas*/
                    donor = useContact;
                else
                {

                    if (donation.PaymentType == (int)PaymentType.SMS || donation.PaymentType == (int)PaymentType.Swish || donation.PaymentType == (int)PaymentType.SMSmanadsdragning)
                    {
                        donor = dataAccess.FindAndUpsertContact(true, donation.DonorInfo.PersonalNo, donation.DonorInfo.FirstName, donation.DonorInfo.LastName, donation.DonorInfo.PrimaryStreetAddressLine1, donation.DonorInfo.PrimaryStreetAddressLine2, donation.DonorInfo.PrimaryPostalCode, donation.DonorInfo.PrimaryCity, donation.DonorInfo.MobilePhone, donation.DonorInfo.PhoneNo, donation.DonorInfo.EmailAddress, donation.SwishName);
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(customerId))
                        {
                            // Hämta individ från customerId om det har skickats in.
                            debugNoteText += "Customer ID: '" + customerId + "'" + Environment.NewLine;
                            donor = dataAccess.FindAndUpsertContact(false, donation.DonorInfo.PersonalNo, donation.DonorInfo.FirstName, donation.DonorInfo.LastName, donation.DonorInfo.PrimaryStreetAddressLine1, donation.DonorInfo.PrimaryStreetAddressLine2, donation.DonorInfo.PrimaryPostalCode, donation.DonorInfo.PrimaryCity, donation.DonorInfo.MobilePhone, donation.DonorInfo.PhoneNo, donation.DonorInfo.EmailAddress, null);
                        }
                        else
                        {
                            donor = dataAccess.GetContactFromCustomerId(customerId);
                        }
                    }
                }

				// REF2016CHANTE 26
//				gift.srk_contactId = new CrmEntityReference(Contact.EntityLogicalName, donor.ContactId.Value);
				gift.srk_contactId = new EntityReference(Contact.EntityLogicalName, donor.ContactId.Value);
				negativeDonor = donor;
            }
            else if (donation.DonationOwnerType == "A")
            {
                debugNoteText += "Anonym gåva" + Environment.NewLine;

                // Anonym gåva från redcross.se. Sök upp och koppla gåvan till "Okänd givare".
                Contact donor = dataAccess.GetUnknownDonorContact();
				// REF2016CHANTE 26
//				gift.srk_contactId = new CrmEntityReference(Contact.EntityLogicalName, donor.ContactId.Value);
				gift.srk_contactId = new EntityReference(Contact.EntityLogicalName, donor.ContactId.Value);
			}
			else if (donation.DonationOwnerType == "F")
            {
                // Hämta företag från customerId om det har skickats in. CustomerId skickas endast med från AX. Från webforms kommer aldrig CustomerId
                // Annars lägg alltid upp nytt företag eftersom OrgNr och namn
                // inte är tillräckligt för säker dubblett(krävs PAR-id också).

                Account donor = null;
                if (string.IsNullOrWhiteSpace(customerId)) //CustomerId kommer aldrig med från webforms
                {
                    donor = dataAccess.InsertAccount(donation.CompanyNo, donation.CompanyName, donation.DonorInfo.FirstName, donation.DonorInfo.LastName, donation.DonorInfo.PrimaryStreetAddressLine1, donation.DonorInfo.PrimaryPostalCode, donation.DonorInfo.PhoneNo, donation.DonorInfo.EmailAddress, donation.Website);

                    debugNoteText += "Nytt konto skapat: " + donor.Name + Environment.NewLine;
                }
                else
                {
                    donor = dataAccess.GetAccountFromAccountId(customerId);
                    debugNoteText += "Befintligt konto funnet: " + donor.Name + Environment.NewLine;
                }

				// REF2016CHANTE 26
//				gift.srk_accountId = new CrmEntityReference(Account.EntityLogicalName, donor.AccountId.Value);
				gift.srk_accountId = new EntityReference(Account.EntityLogicalName, donor.AccountId.Value);
				negativeDonor = donor;

            }
            else if (donation.DonationOwnerType == "O")
            {

                srk_organization organizationDonor = null;
                if (string.IsNullOrWhiteSpace(customerId))
                {
                    organizationDonor = dataAccess.InsertOrganization(donation.CompanyNo, donation.CompanyName, donation.DonorInfo.FirstName, donation.DonorInfo.LastName, donation.DonorInfo.PrimaryStreetAddressLine1, donation.DonorInfo.PrimaryPostalCode, donation.DonorInfo.PhoneNo, donation.DonorInfo.EmailAddress);
                    debugNoteText += "Organisation INTE funnen, ny skapad." + Environment.NewLine;
                }
                else
                {
                    organizationDonor = dataAccess.GetOrganizationFromOrgId(customerId);
                }

				// REF2016CHANTE 26
//				gift.srk_organizationId = new CrmEntityReference(srk_organization.EntityLogicalName, organizationDonor.Id);
				gift.srk_organizationId = new EntityReference(srk_organization.EntityLogicalName, organizationDonor.Id);

			}

			// Koppla till kampanj
			srk_fundraising camp = null;
            if (donation.PaymentType == (int)PaymentType.Swish)
            {
                camp = dataAccess.GetValidCampaignFromSwishNumber(donation.CampaignEventCode, donation.DonationDate);
                debugNoteText += "Swishnummer: " + donation.CampaignEventCode + Environment.NewLine;
            }
            else if (donation.PaymentType == (int)PaymentType.SMS || donation.PaymentType == (int)PaymentType.SMSmanadsdragning)
            {
                if (!string.IsNullOrWhiteSpace(donation.CampaignPrefix))
                {
                    camp = dataAccess.GetValidCampaignFromSMSChannel(donation.CampaignPrefix, donation.CampaignEventCode, donation.DonationDate);
                    debugNoteText += "Prefix: " + donation.CampaignPrefix + " Event: " + donation.CampaignEventCode + Environment.NewLine;
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(donation.CampaignId))
                {
                    camp = dataAccess.GetFundraisingCampaignFromCampaignId(donation.CampaignId);
                    debugNoteText += "CampaignId: " + donation.CampaignId + Environment.NewLine;
                }
            }

            if (camp != null)
            {
				// REF2016CHANTE 26
//				gift.srk_fundraisingId = new CrmEntityReference(srk_fundraising.EntityLogicalName, camp.srk_fundraisingId.Value);
				gift.srk_fundraisingId = new EntityReference(srk_fundraising.EntityLogicalName, camp.srk_fundraisingId.Value);
				debugNoteText += "Kampanj funnen" + Environment.NewLine;
            }
            else
            {
                debugNoteText += "Kampanj INTE funnen" + Environment.NewLine;
            }

            // Spara gåvan
            Guid giftId = dataAccess.Create(gift);
            if (donation.TotalAmount < 0 && negativeDonor != null)
                new CRMDonationHelper().HandleMinusDonation(dataAccess, gift, giftId, negativeDonor);

            dataAccess.CreateNote("Detaljer kring skapandet av gåvan.", debugNoteText, gift);

            return gift;
        }

        /// <summary>
        /// Metod för att spara en autogirobetalbetalning av gåva samt skapa gåva i CRM
        /// </summary>
        /// <param name="bizPayment"></param>
        /// <param name="dataAccess"></param>
        private void CreateDonationFromAutogiroPayment(Payment bizPayment, ICrmDataAccess dataAccess)
        {
            // Validera att vi fått referensen från BizTalk
            if (string.IsNullOrWhiteSpace(bizPayment.PaymentNumber))
                throw new ArgumentException("Referensnummer kan inte vara blankt", "Payment.PaymentNumber");

            // Hämta autogiro från betalningsreferensen.
            srk_autogiro autogiro = null;
            try
            {
                autogiro = dataAccess.GetActiveAutogiroFromReferenceNumber(bizPayment.PaymentNumber);
            }
            catch (Exception e1)
            {
                try
                {
                    autogiro = dataAccess.GetAutogiroFromOldReferenceNumber(bizPayment.PaymentNumber);
                }
                catch (Exception e2)
                {
                    throw new Exception("CRMMainService.CreateDonationFromAutogiroPayment: Inget nytt eller gammalt autogiro kunde hittas för referensnummer: '" + bizPayment.PaymentNumber + "'" + Environment.NewLine +
                                        "Fel angivet vid försök att hitta nytt: " + e1 + Environment.NewLine + Environment.NewLine +
                                        "Fel angivet vid försök att hitta gammalt: " + e2);
                }
            }

            // Skapa och fyll gåvo-objekt
            Donation donationObject = new Donation
            {
                DonationOwnerType = bizPayment.OwnerType,
                DonationDate = bizPayment.PaymentDate, //Samma som betaldatum eftersom inget speciellt gåvodatum finns för autogiro
                PaymentDate = bizPayment.PaymentDate,
                Paid = true,
                DonationType = (int)DonationType.MonthlyGift,
                PaymentType = (int)PaymentType.Autogiro,
                TotalAmount = bizPayment.Amount,
                CampaignId = bizPayment.CampaignId,
                DonorInfo = new Person()
            };
            System.Diagnostics.Debug.WriteLine("CRMTRC: " + "CreateDonation start");

            // Skapa gåva
            srk_donation donation = CreateDonation(donationObject, bizPayment, dataAccess);
            System.Diagnostics.Debug.WriteLine("CRMTRC: " + "CreateDonation start");

			// Koppla gåvan till autogiro
			// REF2016CHANTE 26
//			donation.srk_autogirodonationid = new CrmEntityReference(srk_autogiro.EntityLogicalName, autogiro.srk_autogiroId.Value);
			donation.srk_autogirodonationid = new EntityReference(srk_autogiro.EntityLogicalName, autogiro.srk_autogiroId.Value);
			dataAccess.Update(donation);
        }

        private void CreateDonationFromPayment(Payment payment, ICrmDataAccess dataAccess)
        {
            try
            {
                srk_donation donation = null;
                string logMessage = "Gåvobetalning inkommen." + Environment.NewLine;

                // Kontrollera om det finns någon obetald gåva för angivet OCR-nr. 
                if (!string.IsNullOrWhiteSpace(payment.PaymentNumber))
                {
                    donation = dataAccess.GetUnpayedDonationFromOCRNumber(payment.PaymentNumber);
                }
                //Om ingen gåva med angett OCR-nummer hittades, försök hitta en obetald gåva med hjälp av kampanjkod och individ/konto/organisation
                if (donation == null)
                {
                    donation = dataAccess.GetUnPaidDonationWithoutOCRNumber(payment.OwnerType, payment.CampaignId, payment.CustomerId, payment.PaymentDate, payment.Amount);
                }

                // Om det finns minst en så ska den första markeras som betald. I alla andra fall skapa upp ny gåva.
                if (donation == null)
                {
                    // Skapa och fyll gåvo-objekt
                    Donation donationObject = new Donation()
                    {
                        DonationOwnerType = payment.OwnerType,
                        DonationDate = payment.PaymentDate, // Samma som betaldatum eftersom inget speciellt gåvodatum finns vid inkommen betalning
                        PaymentDate = payment.PaymentDate,
                        Paid = true,
                        DonationType = (int)DonationType.CommonGift,
                        PaymentType = (int)PaymentType.Payment,
                        TotalAmount = payment.Amount,
                        CampaignId = payment.CampaignId,
                        DonorInfo = new Person()
                    };

                    // Skapa gåva
                    donation = CreateDonation(donationObject, payment, dataAccess);
                    logMessage += "Hittade ingen tidigare gåva med OCR nummer: '" + payment.PaymentNumber + "', har därför skapat en ny gåva." + Environment.NewLine;

                    // försök koppla samman gåvan med ett produktköp genom individ och kampanjkod.
                    ConnectDonationToPurchaseIfPossible(donation, dataAccess, ref logMessage);
                }
                else
                {
                    // Markera gåvan som betald.
                    donation.srk_ispayed = true;
                    donation.srk_paymentdate = payment.PaymentDate; //TODO: Kolla med Björn att det är korrekt betaldatum som kommer in.
                    donation.srk_donationdate = payment.PaymentDate; //Gåvodatum ska alltid vara samma som betaldatum. WB43 Release 3.3.
                    donation.srk_voucherno = payment.VoucherNumber;
                    donation.srk_transactionno = payment.TransactionNumber;
					//REF2016CHANTE 14
//					donation.srk_amount = payment.Amount; // Beloppet för betalningen sätts alltid här eftersom man kan ha betalt mer än aviserat belopp.
					donation.srk_amount = new Microsoft.Xrm.Sdk.Money(payment.Amount); // Beloppet för betalningen sätts alltid här eftersom man kan ha betalt mer än aviserat belopp.
					dataAccess.Update(donation);

                    logMessage += "Befintlig gåva hittad med OCR nummer: " + payment.PaymentNumber + ", uppdaterar med betalningsuppgifter." + Environment.NewLine;
                }

                if (donation != null)
                {
                    // om kopplat produktköp, sätt det till betalt och skapa notering på köpet om vad som hänt.
                    if (donation.srk_productgiftid != null)
                    {
                        var purchase = dataAccess.Get<srk_productpurchase>(donation.srk_productgiftid.Id);
                        if (dataAccess.IsProductPurchaseDonationsSumEqualOrLargerToRowSum(purchase, donation))
                        {
                            purchase.srk_paymentdate = payment.PaymentDate;
                            purchase.srk_ispaid = true;
                            dataAccess.Update(purchase);

                            dataAccess.CreateNote("Produktköp betalt", "Betalning inkommen på gåva kopplat till detta produktköp." + Environment.NewLine +
                                              "Betald summa: " + payment.Amount + ", OCR: " + payment.PaymentNumber + ", Kampanj: " + payment.CampaignId,
                                              purchase);
                            logMessage += "Gåvan har ett kopplat produktköp sätter därför det till betalt och skapar en anteckning på produktköpet. " + Environment.NewLine;
                        }
                        else
                        {
                            dataAccess.CreateNote("Produktköp delbetalt", "Betalning inkommen på gåva kopplat till detta produktköp." + Environment.NewLine +
                                              "Betald summa: " + payment.Amount + ", OCR: " + payment.PaymentNumber + ", Kampanj: " + payment.CampaignId,
                                              purchase);
                            logMessage += "Gåvan har ett kopplat produktköp, men gåvosumman är lägre än produkradernas summa. Produktköpet sätts därför ej till betalt. " + Environment.NewLine;
                        }

                    }

                    //inaktivera gåva och skapa notering
                    dataAccess.CreateNote("Gåva betald", logMessage, donation);
                    //dataAccess.SetState(donation, 1, 2); // sätt gåva till Inaktiv - "Stänga" enligt CR57 innebär endast att fälten låses vid betald gåva, vilket göras på alla betlada gåvor. Därför kommenteras denna ut tills vidare.
                }
            }
            catch (Exception e)
            {
                SrkCrmContext.LogOperation(dataAccess, "Error: Crm.MainService.CreateDonationFromPayment", ObjectToString.ToString(payment) + Environment.NewLine + e,
                    CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Error);

                throw e;
            }
        }

        private void ConnectDonationToPurchaseIfPossible(srk_donation donation, ICrmDataAccess dataAccess, ref string LogMessage)
        {
            try
            {
                srk_productpurchase purchase = null;

                if (donation.srk_fundraisingId == null || donation.srk_productgiftid != null)
                    return;

                srk_fundraising campaign = dataAccess.Get<srk_fundraising>(donation.srk_fundraisingId.Id);

                if (donation.srk_contactId != null)
                {
                    string donationCampaignCode = (campaign.srk_campaignid == null ? "" : campaign.srk_campaignid).Trim();

					// Hitta obetalda produktköp kopplade till samma individ med samma kampanjkod, detta täcker in engångsköp av JKE
					// Om vi inte hittar någon, kontrollera om gåvans kampanjkod motsvarar årets kampanjkod för JKE-prenumerationer.
					// I så fall görs uppslag på produktköp med en produkt vars produktnummer motsvarar JKE-prenumerationer.

					//REF2016CHANTE 14
//					purchase = dataAccess.GetUnpaidProductPurchasesFromContactAndCampaign(donation.srk_contactId.Id, donation.srk_fundraisingId.Id, donation.srk_amount);
					purchase = dataAccess.GetUnpaidProductPurchasesFromContactAndCampaign(donation.srk_contactId.Id, donation.srk_fundraisingId.Id, donation.srk_amount?.Value);

					if (purchase != null)
                    {
                        LogMessage += "Produktköp av typen engångsköp hittat med samma kampanjkod kopplat till samma individ." + Environment.NewLine;
                    }
                    else if (purchase == null)
                    {
                        //Logik där prenumerationer slogs upp utifrån samma kampanjkod som på inkommande gåva. Denna logik justeras 
                        //till att jämföra årets jke kampanjkod mot prenumerationer på julklappsetiketter.
                        //purchase = dataAccess.GetSubscriptionFromContactAndCampaign(donation.srk_contactId.Id, donation.srk_fundraisingId.Id);

                        string campaignCodeJKESubscription = dataAccess.GetConfiguration("JKE_Subscription_ThisYearsCampaignCode");

                        if (campaignCodeJKESubscription == donationCampaignCode)
                        {
                            string productIdJKESubscription = dataAccess.GetConfiguration("JKE_Subscription_ProductId");

                            purchase = dataAccess.GetJKESubscriptionForContact(donation.srk_contactId.Id, productIdJKESubscription);

                            if (purchase != null)
                            {
                                LogMessage += "Produktköp av typen prenumeration på JKE som motsvarar årets kampanjkod hittad." + Environment.NewLine;
                            }
                        }
                    }
                }

                if (purchase == null && campaign.srk_product_fundraising_id != null)
                {
                    string campaignCodeJKEProductId = dataAccess.GetConfiguration("JKE_ProductID_ChristmasStickers");
                    srk_product jkeProduct = dataAccess.Get<srk_product>(campaign.srk_product_fundraising_id.Id);
                    //Slå upp kampanjen och hämta produkten
                    if (campaignCodeJKEProductId == (string.IsNullOrWhiteSpace(jkeProduct.srk_itemnumber) ? "" : jkeProduct.srk_itemnumber))
                    {
                        purchase = dataAccess.CreateJKEPurchaseFromOfferDonation(donation, jkeProduct);
                        if (purchase != null)
                        {
                            LogMessage += "Produktköp skapat utifrån gåva på JKE som ej kunde kopplas till befintligt produktköp." + Environment.NewLine;
                        }
                    }
                }

                if (purchase != null)
                {
					// REF2016CHANTE 26
//					donation.srk_productgiftid = new CrmEntityReference(purchase.LogicalName, purchase.Id);
					donation.srk_productgiftid = new EntityReference(purchase.LogicalName, purchase.Id);
					dataAccess.Update(donation);

                    LogMessage += "Gåva sammankopplad med produktköp." + Environment.NewLine;
                }
            }
            catch (Exception e)
            {
                SrkCrmContext.LogOperation(dataAccess,
                    "Error: Crm.MainService.ConnectDonationToPurchaseIfPossible",
                    LogMessage + Environment.NewLine + ObjectToString.ToString(donation) + Environment.NewLine + e,
                    CrmUtility.SystemLogCategory.Integration, CrmUtility.SystemLogType.Error
                );

                throw e;
            }
        }

        private srk_membership CreateMembershipFromRedCross(Membership membership, ICrmDataAccess dataAccess, Contact member)
        {
			// REF2016CHANTE 26
//			CrmEntityReference membershipType;
			EntityReference membershipType;
			decimal? price;
			// REF2016CHANTE 26
//			CrmEntityReference organization = null;
//            CrmEntityReference organizationMemberOrg = null;
//            CrmEntityReference campaign = null;
			EntityReference organization = null;
			EntityReference organizationMemberOrg = null;
			EntityReference campaign = null;
			DateTime? paymentDate = null;
            bool isCampaignCodeOK = false;
            string orderNumber = string.IsNullOrWhiteSpace(membership.OrderNumber) ? "" : membership.OrderNumber;

            // Skapa medlemskap och sätt nödvändig info. OCR-nr sätts inte här utan i plugin PreMembershipCreate.
            srk_membership srkMemShip = new srk_membership();
			// REF2016CHANTE 26
//			var contactId = new CrmEntityReference(Contact.EntityLogicalName, member.ContactId.Value);
			var contactId = new EntityReference(Contact.EntityLogicalName, member.ContactId.Value);

			//int membershipYear = (membership.OrderDate.Month > 9 ? membership.OrderDate.Year + 1 : membership.OrderDate.Year);
			//int membershipYear = membership.OrderDate.Year;
			//startDate = membership.OrderDate;
			//stopDate = new DateTime(membershipYear, 12, 31);
			//validForYear = startDate.Year;

			MembershipDateHelper dateHelper = new MembershipDateHelper(dataAccess);
            if (membership.PaymentType != (int)PaymentType.Avi)
            {
                dateHelper.HandleBreakDate(membership.OrderDate, member.Id);
                if (dateHelper.HandleFutureMemberships(member.Id) == false)
                    throw new CrmRedyException("En medlemskapsbetalning inkom på en individ som redan har max antal framtida medlemskap.");
            }

            var startDate = dateHelper.StartDateNextYear ? new DateTime(dateHelper.RegardingYear, 1, 1) : membership.OrderDate;
            var stopDate = new DateTime(dateHelper.RegardingYear, 12, 31);
            var validForYear = dateHelper.RegardingYear;

            string membershipInternalName = membership.MemberType;
            srk_membershiptype type = dataAccess.GetMembershipType(membershipInternalName);
            if (type == null)
            {
                throw new Exception("Medlemskap kan inte skapas då ingen medlemskapstyp med internt namn '" + membershipInternalName + "' kunde hittas. Kontakta din systemadminstratör.");
            }
            else
            {
				// REF2016CHANTE 26
//				membershipType = new CrmEntityReference(srk_membershiptype.EntityLogicalName, type.srk_membershiptypeId.Value);
                membershipType = new EntityReference(srk_membershiptype.EntityLogicalName, type.srk_membershiptypeId.Value);
				//REF2016CHANTE 14
//				price = type.srk_price;
				price = type.srk_price?.Value;
			}

			// If the member should be connected to a local organization, find the organization from the given postalcode.
			bool failedToGetLocalOrg = false;
            if (membership.LocalOrganization == true)
            {
                srk_organization org = dataAccess.GetOrganizationFromPostalCode(membership.MemberInfo.PrimaryPostalCode, membership.MemberType);

                if (org != null)
                {
					// REF2016CHANTE 26
//					organization = new CrmEntityReference(srk_organization.EntityLogicalName, org.Id);
					organization = new EntityReference(srk_organization.EntityLogicalName, org.Id);
				}
				else
                {
                    failedToGetLocalOrg = true;
                }
            }

            if (membership.LocalOrganization == false || organization == null)
            {
                string redCrossCentralIDForOld = dataAccess.GetConfiguration("ME_RedCrossCentralOrganizationID");
                string redCrossCentralIDForYoung = dataAccess.GetConfiguration("ME_RedCrossCentralOrganizationRKUFID");

                srk_organization centralRCOrganisation;

                if (membership.MemberType == MemberType.MemberRKUF.ToString())
                {
                    centralRCOrganisation = dataAccess.GetOrganizationFromOrgId(redCrossCentralIDForYoung);
                }
                else
                {
                    centralRCOrganisation = dataAccess.GetOrganizationFromOrgId(redCrossCentralIDForOld);
                }

				// REF2016CHANTE 26
//				organization = new CrmEntityReference(srk_organization.EntityLogicalName, centralRCOrganisation.Id);
				organization = new EntityReference(srk_organization.EntityLogicalName, centralRCOrganisation.Id);
			}


			// Sätt Kampanj till värdet som kommer in via BT
			if (!string.IsNullOrEmpty(membership.CampaignCode))
            {
                srk_fundraising fundraising = dataAccess.GetFundraisingCampaignFromCampaignId(membership.CampaignCode);
                if (fundraising != null)
                {
					// REF2016CHANTE 26
//					campaign = new CrmEntityReference(srk_fundraising.EntityLogicalName, fundraising.Id);
					campaign = new EntityReference(srk_fundraising.EntityLogicalName, fundraising.Id);
					isCampaignCodeOK = true;
                }
            }

            // sätt standardkampanj om angiven Kampanj saknas
            if (isCampaignCodeOK == false || campaign == null)
				// REF2016CHANTE 26
//				campaign = new CrmEntityReference(srk_fundraising.EntityLogicalName, dataAccess.GetDefaultMemberCampaign().Id);
				campaign = new EntityReference(srk_fundraising.EntityLogicalName, dataAccess.GetDefaultMemberCampaign().Id);

            var isPaid = membership.Paid;
            //if (isPaid == true)
            //{
            if (membership.PaymentType != (int)PaymentType.Avi)
            {
                paymentDate = membership.PaymentDate; // TODO: ska vi sätta detta även om det inte är betalt?
            }
            //}
            int? paymentType = membership.PaymentType;

            // BORTKOMMENTERAT, DETTA (2 MEDLEMSKAP EFTER 1 OKT) SKA TROLIGTVIS INTE GÖRAS HÄR, UTAN DE SOM ANROPAR FÅR ISF ANROPA DENNA FUNKTION 2 GGR..
            #region notneeded

            // Sätt vilket år som medlemskapet gäller och korrekt slutdatum, beroende på om dagens datum är senare än 1:a oktober.
            //if (membership.OrderDate.Month > 9 && membership.Paid == true)
            //{
            //    // Om månad är större än 9 och det redan är betalt ska medlemskapet gälla till slutet av året, därefter även hela nästa år.
            //    // därför anropas denna metoden igen med ett nytt medlemskap som är för resterande del av detta år 
            //    // och resten av metoden skapar sedan upp det nya medlemskapet för nästkommande år.

            //    Membership tempFreeMemberForRestOfYear = membership;
            //    tempFreeMemberForRestOfYear.PaymentDate = new DateTime(membership.PaymentDate.Value.Year, 9, 25); // date the is slightly before the date that causes this condition
            //    tempFreeMemberForRestOfYear.OrderDate = new DateTime(membership.PaymentDate.Value.Year, 9, 25);
            //    tempFreeMemberForRestOfYear.Paid = true;
            //    srk_membership freeMembership = CreateMembership(tempFreeMemberForRestOfYear, dataAccess, member);

            //    freeMembership.srk_amount = 0;
            //    dataAccess.Update(freeMembership);

            //    // done with free membership, continue with regular membership for rest of year
            //    new FacadeCommon().LogOperation("CrmMainService::CreateMembership, Free membership created rest of year for Contact because of payment date after October 1: ", null, member);
            //}
            // Spara medlemskapet
            //dataAccess.Create(srkMemShip);

            #endregion notneeded


            #region sätt rörelsemedlemsskap om tillämpbart

            var membershipContact = dataAccess.GetContact(contactId);
            var membershipMaximumAgeForYouth = int.Parse(dataAccess.GetConfiguration("MembershipMaximumAgeForYouth"));
            var age = CrmUtility.GetAgeOfContact(membershipContact);
            var isValidForOrganizationMember = (age <= membershipMaximumAgeForYouth);

            //om giltig för rörelsemedlemsskap 
            if (isValidForOrganizationMember)
            {
				// REF2016CHANTE 26
//				var selectedOrganization = dataAccess.GetOrganization(new CrmEntityReference(srk_organization.EntityLogicalName, organization.Id));
				var selectedOrganization = dataAccess.GetOrganization(new EntityReference(srk_organization.EntityLogicalName, organization.Id));

				if (selectedOrganization.srk_organizationtype.Value == (int)MemberType.MemberRKUF)
                {
                    organizationMemberOrg = dataAccess.GetOrganizationFromPostalCode(membership.MemberInfo.PrimaryPostalCode, MemberType.MemberSRK.ToString()).ToEntityReference();
                }
                else
                {
                    organizationMemberOrg = dataAccess.GetOrganizationFromPostalCode(membership.MemberInfo.PrimaryPostalCode, MemberType.MemberRKUF.ToString()).ToEntityReference();
                }
            }
			#endregion

            srk_membership newMembership = CreateMembership
                (
                    contactId,
                    startDate, stopDate, validForYear,
                    membershipType,
                    price ?? 0,
                    price,
                    null,			// Betalt belopp får vi inte från redcross, detta sätts i mottagen betalning när det är bokfört i Ax.
                    organization,
                    organizationMemberOrg,
                    campaign,
                    isPaid,
                    paymentDate,
                    paymentType,
                    "",             // inte betalt så vet ej voucher nummer
                    "",				// inte betalt så vet ej transaktions nummer
                    null,           // från RedCross har vi inga huvudmedlemmar
                    "",             // Kommer in från RedCross & sätts via normala plugins
                    "",             // Kommer in från RedCross & sätts via normala plugins 
                    "",             // Kommer in från RedCross & sätts via normala plugins
                    "",             // Kommer in från RedCross & sätts via normala plugins
                    "",             // Kommer in från RedCross & sätts via normala plugins
                    "",             // Kommer in från RedCross & sätts via normala plugins
                    "",             // Kommer in från RedCross & sätts via normala plugins

                    isValidForOrganizationMember,
                    null,			// Ingen autogirokoppling för medlemskap från redcross
                    orderNumber,
                    false,
                    dataAccess
                );

            dataAccess.CreateNote("Medlemskap skapat", "Medlemskapet är skapat via RedCross.se", newMembership);

            if (failedToGetLocalOrg)
            {
                dataAccess.CreateNote
                    ("Lyckades inte hitta krets.",
                    "Lyckades inte hitta lokal krets för angivet postnummer" + Environment.NewLine + Environment.NewLine +
                    "Logdata: " + Environment.NewLine +
                    ObjectToString.ToString(membership),
                    newMembership
                );
            }

            return newMembership;
        }

        private srk_membership CreateMembership(
			// REF2016CHANTE 26
//			CrmEntityReference contactId,
			EntityReference contactId,
			DateTime startDate, DateTime stopDate, int validForYear,
			// REF2016CHANTE 26
//			CrmEntityReference membershipType,
			EntityReference membershipType,
			decimal price,
            decimal? membershipFee,
            decimal? payedAmount,
			// REF2016CHANTE 26
//			CrmEntityReference organization,
			EntityReference organization,
			// REF2016CHANTE 26
//			CrmEntityReference organizationMembership,
			EntityReference organizationMembership,
			// REF2016CHANTE 26
//			CrmEntityReference campaign,
			EntityReference campaign,
			bool isPaid,
            DateTime? paymentDate,
            int? paymentType,
            string voucherNumber,
            string transactionNumber,
			// REF2016CHANTE 26
//			CrmEntityReference mainMember,
			EntityReference mainMember,
			string recipientName,
            string recipientId,
            string recipientAddressLine1,
            string recipientAddressLine2,
            string recipientPostalCode,
            string recipientCity,
            string recipeintMembershipIncludes,
            //string ocrNumber,
            bool isOrganizationMember,
			// REF2016CHANTE 26
//			CrmEntityReference autogiro,
			EntityReference autogiro,
			string orderNumber,
            bool aviPrint,
            ICrmDataAccess dataAccess)
        {
            srk_membership srkMemShip = new srk_membership();
			// REF2016CHANTE 26
//			srkMemShip.srk_contactId = new CrmEntityReference(Contact.EntityLogicalName, contactId.Id);
			srkMemShip.srk_contactId = new EntityReference(Contact.EntityLogicalName, contactId.Id);
			srkMemShip.srk_startdate = startDate;
            srkMemShip.srk_stopdate = stopDate;
            srkMemShip.srk_regardingyear = validForYear.ToString();
            srkMemShip.srk_IsOrganizationMember = isOrganizationMember;
			// REF2016CHANTE 26
//			srkMemShip.srk_membershiptypeId = new CrmEntityReference(srk_membershiptype.EntityLogicalName, membershipType.Id);
			srkMemShip.srk_membershiptypeId = new EntityReference(srk_membershiptype.EntityLogicalName, membershipType.Id);
			//REF2016CHANTE 14
//			srkMemShip.srk_membership_fee = membershipFee;
			srkMemShip.srk_membership_fee = (membershipFee != null ? new Microsoft.Xrm.Sdk.Money(membershipFee.Value) : (Microsoft.Xrm.Sdk.Money)null);
			//REF2016CHANTE 14
//			srkMemShip.srk_amount = price;
			srkMemShip.srk_amount = new Microsoft.Xrm.Sdk.Money(price);
			//REF2016CHANTE 14
//			srkMemShip.srk_payed_amount = (isPaid ? payedAmount : null);
			srkMemShip.srk_payed_amount = (isPaid ? (payedAmount != null ? new Microsoft.Xrm.Sdk.Money(payedAmount.Value) : (Microsoft.Xrm.Sdk.Money)null) : (Microsoft.Xrm.Sdk.Money)null);
			//REF2016CHANTE 14
//			srkMemShip.srk_distributed_amount = (isPaid ? membershipFee : null);
			srkMemShip.srk_distributed_amount = (isPaid ? (membershipFee != null ? new Microsoft.Xrm.Sdk.Money(membershipFee.Value) : (Microsoft.Xrm.Sdk.Money)null) : (Microsoft.Xrm.Sdk.Money)null);
			// REF2016CHANTE 26
//			srkMemShip.srk_organizationId = new CrmEntityReference(srk_organization.EntityLogicalName, organization.Id);
			srkMemShip.srk_organizationId = new EntityReference(srk_organization.EntityLogicalName, organization.Id);
			srkMemShip.srk_ordernumber = orderNumber;
            srkMemShip.srk__avi_print = aviPrint;

            //organisation för rörelsemedlemsskap
            if (organizationMembership != null)
				// REF2016CHANTE 26
//				srkMemShip.srk_OrganizationMembership_OrgId = new CrmEntityReference(srk_organization.EntityLogicalName, organizationMembership.Id);
                srkMemShip.srk_OrganizationMembership_OrgId = new EntityReference(srk_organization.EntityLogicalName, organizationMembership.Id);

            srkMemShip.srk_kampanjkod_id = campaign;

            srkMemShip.srk_ispayed = isPaid;

            // if (isPaid)
            //{
            srkMemShip.srk_paymentdate = paymentDate;
            //}

            //Om medlemskapet ska kopplas till autogiro
            if (autogiro != null)
            {
                paymentType = (int)PaymentType.Autogiro;
				// REF2016CHANTE 26
//				srkMemShip.srk_autogirouppdragid = new CrmEntityReference(autogiro.LogicalName, autogiro.Id);
				srkMemShip.srk_autogirouppdragid = new EntityReference(autogiro.LogicalName, autogiro.Id);

			}

			//REF2016CHANTE 14
//			srkMemShip.srk_paymenttype = (int)paymentType;
			srkMemShip.srk_paymenttype = new Microsoft.Xrm.Sdk.OptionSetValue((int)paymentType);

			if (voucherNumber != null && voucherNumber.Length > 0)
            {
                srkMemShip.srk_voucherno = voucherNumber;
            }

            if (transactionNumber != null && transactionNumber.Length > 0)
            {
                srkMemShip.srk_transactionno = transactionNumber;
            }

            if (mainMember != null)
            {
                Guid mainMemberId = dataAccess.GetMembershipForYear(mainMember.Id, validForYear.ToString());
                if (mainMemberId != Guid.Empty)
					// REF2016CHANTE 26
//					srkMemShip.srk_mainmemberId = new CrmEntityReference(srk_membership.EntityLogicalName, mainMemberId);
                    srkMemShip.srk_mainmemberId = new EntityReference(srk_membership.EntityLogicalName, mainMemberId);
            }

            srkMemShip.srk_recipient_name = string.IsNullOrWhiteSpace(recipientName) ? "" : recipientName;
            srkMemShip.srk_recipient_contactid = string.IsNullOrWhiteSpace(recipientId) ? "" : recipientId;
            srkMemShip.srk_recipient_addressline1 = string.IsNullOrWhiteSpace(recipientAddressLine1) ? "" : recipientAddressLine1;
            srkMemShip.srk_recipient_addressline2 = string.IsNullOrWhiteSpace(recipientAddressLine2) ? "" : recipientAddressLine2;
            srkMemShip.srk_recipient_postalcode = string.IsNullOrWhiteSpace(recipientPostalCode) ? "" : recipientPostalCode;
            srkMemShip.srk_recipient_city = string.IsNullOrWhiteSpace(recipientCity) ? "" : recipientCity;
            srkMemShip.srk_membership_includes = string.IsNullOrWhiteSpace(recipeintMembershipIncludes) ? "" : recipeintMembershipIncludes;

            dataAccess.Create(srkMemShip);

            return srkMemShip;
        }

		// REF2016CHANTE 26
//		private CrmEntityReference GetOrganization(Membership membership, ICrmDataAccess dataAccess)
        private EntityReference GetOrganization(Membership membership, ICrmDataAccess dataAccess)
        {
			// REF2016CHANTE 26
//			CrmEntityReference organization = null;
			EntityReference organization = null;

			// If the member should be connected to a local organization, find the organization from the given postalcode.            
			if (membership.LocalOrganization == true)
            {
                srk_organization org = dataAccess.GetOrganizationFromPostalCode(membership.MemberInfo.PrimaryPostalCode, membership.MemberType);

                if (org != null)
                {
					// REF2016CHANTE 26
//					organization = new CrmEntityReference(srk_organization.EntityLogicalName, org.Id);
					organization = new EntityReference(srk_organization.EntityLogicalName, org.Id);
				}
			}

            if (membership.LocalOrganization == false || organization == null)
            {
                string redCrossCentralIDForOld = dataAccess.GetConfiguration("ME_RedCrossCentralOrganizationID");
                string redCrossCentralIDForYoung = dataAccess.GetConfiguration("ME_RedCrossCentralOrganizationRKUFID");

                srk_organization centralRCOrganisation;

                if (membership.MemberType == MemberType.MemberRKUF.ToString())
                {
                    centralRCOrganisation = dataAccess.GetOrganizationFromOrgId(redCrossCentralIDForYoung);
                }
                else
                {
                    centralRCOrganisation = dataAccess.GetOrganizationFromOrgId(redCrossCentralIDForOld);
                }

				// REF2016CHANTE 26
//				organization = new CrmEntityReference(srk_organization.EntityLogicalName, centralRCOrganisation.Id);
				organization = new EntityReference(srk_organization.EntityLogicalName, centralRCOrganisation.Id);
			}

			return organization;

        }

        /// <summary>
        /// Validate that a donation was made on a campaign code that was active on the the date of the donation.
        /// Returns a string containing an error message if validation fails, or null if validation succeeds
        /// </summary>
        private string ValidateCampaignCodeForDonation(srk_fundraising camp, DateTime donationDate)
        {
            if (camp == null)
                return "Det gick inte att hämta en kampanjkod för gåvan.";

			//REF2016CHANTE 14
//			if (camp.srk_campaignstatus == (int)CampaignStatus.Preliminary)
            if (camp.srk_campaignstatus?.Value == (int)CampaignStatus.Preliminary)
				return String.Format("Kampanjkod {0} har status \"Under upprättande\", inga gåvor kan skapas på kampanjen.", camp.srk_campaignid);

            DateTime campStartdate = ((DateTime)camp.srk_startdate).ToLocalTime();
            DateTime campStopdate = camp.srk_stopdate != null ? ((DateTime)camp.srk_stopdate).ToLocalTime() : DateTime.MaxValue;

            if (campStartdate > donationDate.Date || campStopdate < donationDate.Date)
                return String.Format("Gåvans gåvodatum {0} infaller inte mellan startdatum ({1}) och slutdatum ({2}) för kampanjkod {3}.", donationDate.ToString("d"), campStartdate.ToString("d"), campStopdate.ToString("d"), camp.srk_campaignid);

            return null;//No error
        }

        private Common.SystemJobType TranslateJobType(SystemJobType facadeType)
        {
            switch (facadeType)
            {
                case SystemJobType.MemberFeeDistribution:
                    return Common.SystemJobType.MemberFeeDistribution;
                case SystemJobType.PostalCodeImport:
                    return Common.SystemJobType.PostalCodeImport;
                case SystemJobType.SPARExport:
                    return Common.SystemJobType.SPARExport;
                case SystemJobType.HittaExport:
                    return Common.SystemJobType.HittaExport;
                case SystemJobType.SPARImport:
                    return Common.SystemJobType.SPARImport;
                case SystemJobType.HittaImport:
                    return Common.SystemJobType.HittaImport;
                case SystemJobType.WelcomeLetter:
                    return Common.SystemJobType.WelcomeLetter;
                case SystemJobType.YearlyMemberNotification:
                    return Common.SystemJobType.YearlyMemberNotification;
                case SystemJobType.ClumpSMSExport:
                    return Common.SystemJobType.ClumpSMSExport;
                case SystemJobType.AutogiroImport:
                    return Common.SystemJobType.AutogiroImport;
                case SystemJobType.ClumpBookingExport:
                    return Common.SystemJobType.ClumpBookingExport;
                case SystemJobType.AutogiroSubscriptions:
                    return Common.SystemJobType.AutogiroSubscriptions;
            }
            return Common.SystemJobType.Unknown;
        }

        #endregion
    }
}
