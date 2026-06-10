export interface RegistrationResponse {
    data: any;
    isSuccess: boolean;
    errorMessage: string;
    errorCode: string;
    timestamp: string;
    metadata: any;
}



//Verify the code (nonce) and get the actual token
export const VerifyRegistrationNonce = async (_nonce: string): Promise<RegistrationResponse> => {
     const url = `https://dragon.bookingbuilder.com/AutofillService/api/Registration/get-client-registration-from-nonce?nonce=${encodeURIComponent(_nonce)}`;
    // const url = "https://dragon.bookingbuilder.com/AutofillService/api/Registration/get-client-registration-from-nonce?nonce=P1LYzF1Luk2wBqkJCJm25BExzWX0-CBUym-x";

    const response = await fetch(url, { method: "GET", redirect: "follow" });
    if (!response.ok) throw new Error("Verification request failed");
    return await response.json();
};