import { Client } from "@microsoft/microsoft-graph-client";

function createGraphClient(accessToken: string) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

function toGraphLocalDateTime(date: Date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const h = String(date.getHours()).padStart(2, "0");
  const min = String(date.getMinutes()).padStart(2, "0");
  const s = String(date.getSeconds()).padStart(2, "0");

  return `${y}-${m}-${d}T${h}:${min}:${s}`;
}

export type GraphUserProfile = {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
};

function escapeSearchText(value: string) {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .trim();
}

function escapeODataString(value: string) {
  return value.replace(/'/g, "''");
}

export async function searchUsers(
  accessToken: string,
  keyword: string
): Promise<GraphUserProfile[]> {
  const client = createGraphClient(accessToken);

  const searchText = escapeSearchText(keyword);

  if (!searchText) {
    return [];
  }

  const result = await client
    .api("/users")
    .header("ConsistencyLevel", "eventual")
    .query({
      $search:
        `"displayName:${searchText}" OR ` +
        `"mail:${searchText}" OR ` +
        `"userPrincipalName:${searchText}"`,
      $select: "displayName,mail,userPrincipalName",
      $top: "8",
      $count: "true",
    })
    .get();

  return result.value || [];
}

export async function getMyEvents(
  accessToken: string,
  startDate: Date,
  endDate: Date
) {
  const client = createGraphClient(accessToken);

  const result = await client
    .api("/me/calendar/calendarView")
    .header("Prefer", 'outlook.timezone="Tokyo Standard Time"')
    .query({
      startDateTime: startDate.toISOString(),
      endDateTime: endDate.toISOString(),
    })
    .orderby("start/dateTime")
    .get();

  return result.value;
}

export async function getScheduleForUser(
  accessToken: string,
  userPrincipalName: string,
  startDate: Date,
  endDate: Date
) {
  const client = createGraphClient(accessToken);

  const result = await client
    .api("/me/calendar/getSchedule")
    .header("Prefer", 'outlook.timezone="Tokyo Standard Time"')
    .post({
      schedules: [userPrincipalName],
      startTime: {
        dateTime: toGraphLocalDateTime(startDate),
        timeZone: "Tokyo Standard Time",
      },
      endTime: {
        dateTime: toGraphLocalDateTime(endDate),
        timeZone: "Tokyo Standard Time",
      },
      availabilityViewInterval: 30,
    });

  const scheduleInfo = result.value?.[0];

  if (!scheduleInfo) {
    return [];
  }

  if (scheduleInfo.error) {
    throw new Error(
      scheduleInfo.error.message ||
        "getSchedule failed"
    );
  }

  return scheduleInfo.scheduleItems || [];
}

export async function getUserProfile(
  accessToken: string,
  mailOrUserPrincipalName: string
) {
  const client = createGraphClient(accessToken);

  const value = escapeODataString(
    mailOrUserPrincipalName
  );

  const result = await client
    .api("/users")
    .filter(
      `mail eq '${value}' or userPrincipalName eq '${value}'`
    )
    .select(
      "displayName,mail,userPrincipalName"
    )
    .get();

  if (
    !result.value ||
    result.value.length === 0
  ) {
    throw new Error(
      "User not found"
    );
  }

  return result.value[0];
}