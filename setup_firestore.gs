function setupFirestoreProperties() {
  const props = PropertiesService.getScriptProperties();
  
  // Set the Firestore configuration
  props.setProperty('FIRESTORE_PROJECT_ID', 'tallyappdb');
  props.setProperty('FIRESTORE_SERVICE_ACCOUNT_JSON', JSON.stringify({
    "type": "service_account",
    "project_id": "tallyappdb",
    "private_key_id": "c7ad0daf1752204a89b8c0258ad2d3eca5f7e74b",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCPe0A3NiXnwo2b\nDY8QO4m+X2ZC6zYZPRMlG111cMic3yxh4O6NX3gC8QCf6jAebSPYd23RecbSdafo\nnPr6YykgQIq7YFUD2cF3j75dvl3vCusAghzitGInHa2QWnf3L7sTBJ4FfjRAzbcx\nfKxuE/6Rk6wo72Nl/lN5jSqc1lR48TFRqW2uFDjr8ZQm4KGb/YNw0LAGdgF03Nx+\ndgGhfgH000rxFvo1zr7G3PS5zMauzIMwQNsMorHN5ZfQtdGzP9zh7UplM6gC6Vt3\nYIG3K98Xc6gY8ldfrPSQbjCGQkz3ctfo7PV0N+txD5ou4612SoR6GFMSZmXYB+Gp\nylmH7dfJAgMBAAECggEACEYj/ZsUhAEi39hgEeBREOaY885zC/EK79acZNeYZf1z\npoBNWV8yQT8rGU2ssTB7gvHeJhiG59+Li30z9ujtd/Po3CXRuTtfLfke0fKWoaCI\nTTrzlfovd9B9EAg7h0hcGhxnzWrJ8hu6zvKA0gc6pKP2p1Sor3rDTwzeMnWkoOW3\nnkt84kXi4fNtWk5fAqIExifbjnGQlXI8C45T9p4ccjy/CF+WwUgnBHW1bD2PhnO8\n0HfIRUF67j+cC9wDv+8xnueZD+mF1jGdjNR9hQxCzo+h/s9iofOMcJ69/JZrgDUB\nxhhEcQilQWHu8RIRGx9sP6TV/Xej6tGbM8CS7mJb5QKBgQDKjf967DgvvuszeW0R\n6VEM5Rdf6Wm8QOUHMFGWloNt6LW7/OXXMo0S5lX6X9p4aafcfxI+QrKtNx4XjHPh\nYEIbtNEkvYv/p3T2Qh5h6b8/RdntpUX8Dgx/BJoW/twmHCx6w+eu5rQSyl4gNxOf\nC6rV+n03IHC6xTBRI9OoxqtVVQKBgQC1VwV3sWXSa6UaD73Qn/WCGa1xwk3j1BR0\nB8D6HCDlnVEBGgymp2jfIWNoBesS2IYgHbWQnDspM/acKPR0p4hPQCnZctjeni4e\nBM5EnZlXZtCtXYOVBql+HuBEuT5q0sXu7XBECwLjB7cx/uy278LuY1pxoWjzIgw2\ndTasfYB4pQKBgCXp3Fc63e158/ZnlRXUMTqTqwNl4ke9ceEdlzfzplgtRc8RMEhO\niv9to8xAotkFFcl884I5483uUjCgNpSJ+vXKq5cP76ODya3KhHtZUFXckkUcOXRL\ngOVJjDofrD9p6J12C+i+ABjATkwDYyXS/kAdKdDmvpMHE9ssNJMT1MAdAoGAFvXf\nwVn61HvAud4IYKQAR+rMZTz+87vkoxRgQMrS93/Fw+fydh94TdhCZ+dn7kjwIpzg\nYmEFtJ/Zt8gPu72CqtYq9lduUjGhETf2jgWOZtyjy2+tu27OgyORreHun0P0v93c\neyERyoEtluz6N+d/vQNjbkPLsA09kp+gvpJdc+kCgYBbLx9HpG9bDVuTPCYY3vZ1\nK5cOz03XB6x7e/cHfPabKjMdTvD8mBaiFcBvgx0C8ZBJV6bYGTsk9rQZcIKJteCi\nHw3Jl/asHl7K9lgDWMijv8nwrxApOzCwlv/aVvZSlArPAXBYFHuEdmgMhOxtRupW\nyKo5K4olUNhAttbMFtlYwQ==\n-----END PRIVATE KEY-----\n",
    "client_email": "firebase-adminsdk-fbsvc@tallyappdb.iam.gserviceaccount.com",
    "client_id": "113909597226496730139",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/firebase-adminsdk-fbsvc%40tallyappdb.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
  }));
  
  console.log('Firestore credentials have been set in Script Properties.');
}
