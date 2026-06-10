
export interface ClientInfo {
  emailClient: string;
  emailClientVersion: string;
  operatingSystem: string;
}

export const getClientInfo = (): ClientInfo => {
  const diagnostics = Office?.context?.mailbox?.diagnostics;

  const emailClient = diagnostics?.hostName || "Outlook";

  const emailClientVersion =
    diagnostics?.hostVersion || "Unknown";

  const ua = navigator.userAgent;

  let operatingSystem = "Unknown";

  if (ua.includes("Windows")) {
    operatingSystem = "Windows";
  } else if (ua.includes("Mac")) {
    operatingSystem = "macOS";
  } else if (ua.includes("Linux")) {
    operatingSystem = "Linux";
  } else if (ua.includes("Android")) {
    operatingSystem = "Android";
  } else if (
    ua.includes("iPhone") ||
    ua.includes("iPad")
  ) {
    operatingSystem = "iOS";
  }

  return {
    emailClient,
    emailClientVersion,
    operatingSystem,
  };
};