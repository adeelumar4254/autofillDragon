export const reportErrorApi = async (
  uniqueId: string,
  description: string
) => {
  try {
    const token = localStorage.getItem("token");
const apiUrl = process.env.Report_Error; 
    const response = await fetch(
      apiUrl,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
        },
        body: JSON.stringify({
          uniqueId,
          description,
        }),
      }
    );

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    return await response.text();
  } catch (err) {
    console.log("Fetch failed:", err);
    throw err;
  }
};