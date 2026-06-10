export interface RegistrationResponse {
    data: string;
    isSuccess: boolean;
    errorMessage: string;
    errorCode: string;
    timestamp: string;
    metadata: any;
}

export const GetRegistrationEmail = async (email: string): Promise<RegistrationResponse> => {
    const url = `https://dragon.bookingbuilder.com/AutofillService/api/Registration/email-addin-submit-email?email=${encodeURIComponent(email)}&emailType=outlook`;
    
    const requestOptions = {
        method: "GET",
        redirect: "follow" as RequestRedirect
    };

    const response = await fetch(url, requestOptions);
    if (!response.ok) {
        throw new Error("Network response was not ok");
    }
    return await response.json();
};