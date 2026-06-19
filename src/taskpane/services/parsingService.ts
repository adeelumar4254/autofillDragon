
import { getClientInfo } from "../utils/clientInfo";
const  apiUrl =  process.env.PARSING_API 
  
export const parseEmail = async (emailMessage: any) => {
  const token = localStorage.getItem("token");

  if (!token) {
    throw new Error("Token not found");
  }

  const { emailClient, emailClientVersion, operatingSystem } = getClientInfo();
console.log(emailMessage);
  const response = await fetch(apiUrl, {
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
