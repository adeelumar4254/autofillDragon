export interface RegistrationResponse {
    data: string;
    isSuccess: boolean;
    errorMessage: string;
    errorCode: string;
    timestamp: string;
    metadata: any;
}

export const GetRegistrationEmail = async (email: string): Promise<RegistrationResponse> => {
    const apiUrl = `${process.env.Register_Email}email=${encodeURIComponent(email)}&emailType=outlook`;
    
    const requestOptions = {
        method: "GET",
        redirect: "follow" as RequestRedirect
    };

    const response = await fetch(apiUrl, requestOptions);
    if (!response.ok) {
        throw new Error("Network response was not ok");
    }
    return await response.json();
};