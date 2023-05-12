// JavaScript code
const form = document.getElementById("myForm");
const recipientEmail = "recipient@example.com"; // Change this to your email address

form.addEventListener("submit", function(event) {
  event.preventDefault(); // prevent the form from submitting normally
  
  // Get form input values
  const name = document.getElementById("name").value;
  const email = document.getElementById("email").value;
  const message = document.getElementById("message").value;
  
  // Organize input values
  const emailBody = `Name: ${name}\nEmail: ${email}\nMessage: ${message}`;
  
  // Send email
  const emailEndpoint = "https://graph.microsoft.com/v1.0/me/sendMail";
  const accessToken = "<your-access-token>";
  
  fetch(emailEndpoint, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      message: {
        subject: "New message from your website",
        body: {
          contentType: "Text",
          content: emailBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipientEmail
            }
          }
        ]
      },
      saveToSentItems: "true"
    })
  })
  .then(response => {
    if (response.ok) {
      // Show alert
      alert("Your message has been sent.");
    } else {
      throw new Error("Error sending email.");
    }
  })
  .catch(error => {
    console.error(error);
  });
  
  // Clear form input values
  form.reset();
});
