SELECT  CONTACTS.phone, CONTACTS.fax, CONTACTS.cellular, CONTACTS.email, 
      CONTACTS.messanger_name AS duty, CONTACTS.CONTACT_NAME, COMPANIES.COMPANY_NAME, 
      CONTACTS.CONTACT_ID,CONTACTS.contact_address,cONTACTS.contact_city_Name, CONTACTS.COMPANY_ID,contact_type=dbo.get_TypeName(CONTACTS.CONTACT_ID)   FROM CONTACTS LEFT OUTER JOIN
      COMPANIES ON CONTACTS.COMPANY_ID = COMPANIES.COMPANY_ID WHERE CONTACTS.ORGANIZATION_ID = 8