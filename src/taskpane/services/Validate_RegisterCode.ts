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
     const apiUrl =`${process.env.Register_Nonce}nonce=${encodeURIComponent(_nonce)}`;

    const response = await fetch(apiUrl,
         { method: "GET",
             redirect: "follow" }
        );
    if (!response.ok) throw new Error("Verification request failed");
    return await response.json();
};