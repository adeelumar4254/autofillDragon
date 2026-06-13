
import { getClientInfo } from "../utils/clientInfo";
const  API_URL =  process.env.PARSING_API 
 //const API_URL =  "https://bb-admin-dev-aiparsing-h6egbrchhad2gcc5.eastus-01.azurewebsites.net/Parsing/api/ParsingService/email-parsing";
  
export const parseEmail = async (emailMessage: any) => {
  const token = localStorage.getItem("token");

  if (!token) {
    throw new Error("Token not found");
  }

  const { emailClient, emailClientVersion, operatingSystem } = getClientInfo();
console.log(emailMessage);
  const response = await fetch(API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
    },
    body: JSON.stringify({
      emailMessage: { 
        Attachments:emailMessage.Attachments,
        Recipients:emailMessage.Recipients,
        Sender:emailMessage.Sender,
        ReceivedBy:emailMessage.ReceivedBy,
        Subject: emailMessage.Subject,
        BodyText: emailMessage.BodyText,
        BodyHtml: emailMessage.BodyHtml,
    
      },
          
       emailClient: emailClient,         
       emailClientVersion: emailClientVersion,  
       operatingSystem: operatingSystem,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(errorText);
  }
console.log("Parsing API Response:", await response.clone().json());
  return response.json();
};
