export const msalConfig = {
  auth: {
    clientId: "51913406-3d5a-44bc-847c-48f21f21c623",
    authority:
      "https://login.microsoftonline.com/e9b67710-82df-42cc-9329-96d07c440be2",

    redirectUri:
  import.meta.env.DEV
    ? "http://localhost:5173"
    : "https://ditishikawakenji.github.io/TeamScheduleBoard/",

    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: [
    "User.Read",
    "Calendars.Read",
    "Calendars.Read.Shared",
    "User.ReadBasic.All",
  ],
};