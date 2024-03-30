function myfunction(){
    console.log(typeof(Math.random().toString(36)));
  }
  
  function listAllUsers() {
    let pageToken;
    let page;
    do {
      page = AdminDirectory.Users.list({
        domain: 'etegravata.com.br',
        orderBy: 'givenName',
        maxResults: 1,
        pageToken: pageToken
      });
      const users = page.users;
      if (!users) {
        console.log('No users found.');
        return;
      }
      // Print the user's full name and email.
      for (const user of users) {
        console.log('%s (%s)', user.name.fullName, user);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);
  }
  
  /*  
  {
    "isAdmin": false,
    "nonEditableAliases": ["modeloteste@etegravata.com.br.test-google-a.com"],
    "lastLoginTime": "1970-01-01T00:00:00.000Z",
    "isMailboxSetup": true,
    "ipWhitelisted": false,
    "externalIds": [
      {
        "value": "aquiVaioCPF",
        "type": "organization"
      }
    ],
    "id": "111907347277341027897",
    "isEnrolledIn2Sv": false,
    "isDelegatedAdmin": false,
    "agreedToTerms": false,
    "languages": [
      {
        "languageCode": "pt",
        "preference": "preferred"
      }
    ],
    "phones": [
      {
        "type": "work",
        "value": "123456789"
      }
    ],
    "etag": "\"WK9tybb2TgVN-FukjV5Q7sjTXQHDeuZZ70913_bjogQ/yvZ_MGb0jdBkn8PIf7XWeCzOWEA\"",
    "primaryEmail": "modeloteste@etegravata.com.br",
    "suspended": false,
    "name": {
      "fullName": "Teste Modelo Tito",
      "familyName": "Tito",
      "givenName": "Teste Modelo"
    },
    "customerId": "C02zyp7pa",
    "changePasswordAtNextLogin": true,
    "creationTime": "2024-03-25T19:05:37.000Z",
    "isEnforcedIn2Sv": false,
    "kind": "admin#directory#user",
    "archived": false,
    "includeInGlobalAddressList": true,
    "emails": [
      {
        "type": "work",
        "address": "Esseehemailsecundario@gmail.com"
      },
      {
        "primary": true,
        "address": "modeloteste@etegravata.com.br"
      },
      {
        "address": "modeloteste@etegravata.com.br.test-google-a.com"
      }
    ],
    "orgUnitPath": "/TITO"
  }
  */  