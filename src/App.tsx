import { useMsal } from "@azure/msal-react";
import {
  useEffect,
  useMemo,
  useState,
} from "react";
import type { CSSProperties } from "react";
import { loginRequest } from "./auth/msalConfig";
import {
  getMyEvents,
  getScheduleForUser,
  getUserProfile,
  searchUsers,
} from "./services/graphService";

type DisplayUser = {
  email: string;
  displayName: string;
};

type UserCandidate = {
  email: string;
  displayName: string;
  userPrincipalName: string;
};

type CalendarEvent = {
  id: string;
  userEmail: string;
  subject?: string;
  location?: string;
  start?: {
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    dateTime?: string;
    timeZone?: string;
  };
  showAs?: string;
  status?: string;
  isError?: boolean;
};

type EventsByUser = {
  [userEmail: string]: CalendarEvent[];
};

declare const __BUILD_TIME__: string;

function App() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const myEmail =
    account?.username || "";

  const formatDisplayName = (
    name: string
  ) => {
    return name.split("/")[0].trim();
  };

  const myDisplayName =
    formatDisplayName(
      account?.name ||
        account?.username ||
        "自分"
    );

  const [eventsByUser, setEventsByUser] =
    useState<EventsByUser>({});

  const [displayUsers, setDisplayUsers] =
    useState<DisplayUser[]>([]);

  const [emailInput, setEmailInput] =
    useState("");

  const [
  userCandidates,
  setUserCandidates,
] = useState<UserCandidate[]>([]);

const [
  showUserCandidates,
  setShowUserCandidates,
] = useState(false);

const [
  isSearchingUsers,
  setIsSearchingUsers,
] = useState(false);

  const [message, setMessage] =
    useState("");

  const [loadingSchedule, setLoadingSchedule] =
  useState(true);

  const [weekOffset, setWeekOffset] =
    useState(0);

  const [wrapText, setWrapText] =
    useState(false);

  const [showLocation, setShowLocation] =
    useState(true);

  const [showSidebar, setShowSidebar] =
    useState(true);

  const [initialized, setInitialized] =
    useState(false);

  const buttonStyle: CSSProperties = {
    border: "1px solid #cbd5e1",
    borderRadius: 4,
    background: "#ffffff",
    padding: "3px 8px",
    cursor: "pointer",
    fontSize: 12,
    lineHeight: 1.4,
  };
const primaryButtonStyle: CSSProperties = {
  ...buttonStyle,
  background: "#2563eb",
  color: "#ffffff",
  border: "1px solid #2563eb",
  fontWeight: 600,
};
 useEffect(() => {
  const saved = localStorage.getItem(
    "teamScheduleSettings"
  );

  if (!saved) {
    setInitialized(true);
    return;
  }

  const settings = JSON.parse(saved);

  const savedUsers =
    settings.displayUsers || [];

  const normalizedUsers: DisplayUser[] =
    savedUsers
      .map((user: any) => {
        if (typeof user === "string") {
          return {
            email: user,
            displayName: user,
          };
        }

        return {
          email: user.email,
          displayName:
            user.displayName ||
            user.name ||
            user.email,
        };
      })
      .filter(
        (user: DisplayUser) =>
          user.email
      );

  setDisplayUsers(normalizedUsers);
  setShowSidebar(
    settings.showSidebar ?? true
  );
  setWrapText(
    settings.wrapText ?? false
  );
  setShowLocation(
    settings.showLocation ?? true
  );

  setInitialized(true);
}, []);

const saveSettings = (
  users: DisplayUser[],
  sidebar = showSidebar,
  wrap = wrapText,
  location = showLocation
) => {
  localStorage.setItem(
    "teamScheduleSettings",
    JSON.stringify({
      displayUsers: users,
      showSidebar: sidebar,
      wrapText: wrap,
      showLocation: location,
    })
  );
};

  const signIn = async () => {
    await instance.loginRedirect(
      loginRequest
    );
  };

  const signOut = async () => {
    await instance.logoutRedirect();
  };

  const getAccessToken = async () => {
    if (!account) {
      throw new Error(
        "ログインアカウントが取得できません。"
      );
    }

    const token =
      await instance.acquireTokenSilent({
        ...loginRequest,
        account,
      });

    return token.accessToken;
  };

  useEffect(() => {
    const keyword =
      emailInput.trim();

if (
  !account ||
  keyword.length < 2 ||
  keyword.includes("@")
) {
  setUserCandidates([]);
  setShowUserCandidates(false);
  return;
}
    let canceled = false;

    const timer =
      window.setTimeout(
        async () => {
          try {
  setIsSearchingUsers(true);
  setShowUserCandidates(true);

  const token =
    await instance.acquireTokenSilent({
                ...loginRequest,
                account,
              });

            const users =
              await searchUsers(
                token.accessToken,
                keyword
              );

            if (canceled) {
              return;
            }

            const candidates =
              users
                .map((user) => {
                  const email =
                    user.mail ||
                    user.userPrincipalName ||
                    "";

                  return {
                    email:
                      email.toLowerCase(),
                    userPrincipalName:
                      user.userPrincipalName ||
                      email,
                    displayName:
                      formatDisplayName(
                        user.displayName ||
                          email
                      ),
                  };
                })
                .filter(
                  (user) =>
                    user.email &&
                    user.email !==
                      myEmail.toLowerCase() &&
                    !displayUsers.some(
                      (displayUser) =>
                        displayUser.email ===
                        user.email
                    )
                );

setUserCandidates(
  candidates
);

setShowUserCandidates(true);
          } catch (error) {
            console.error(error);

            if (!canceled) {
              setUserCandidates([]);
              setShowUserCandidates(false);
            }
          } finally {
            if (!canceled) {
              setIsSearchingUsers(false);
            }
          }
        },
        300
      );

    return () => {
      canceled = true;
      window.clearTimeout(timer);
    };
  }, [
    emailInput,
    account,
    instance,
    displayUsers,
    myEmail,
  ]);

  const selectUserCandidate = (
    candidate: UserCandidate
  ) => {
    setEmailInput(candidate.email);
    setShowUserCandidates(false);
    setMessage(
      `${candidate.displayName} を選択しました。`
    );
  };

  const addUserByEmail = async () => {
    const email =
      emailInput.trim().toLowerCase();

    if (!email) {
      setMessage(
        "メールアドレスを入力してください。"
      );
      return;
    }

    if (!email.includes("@")) {
      setMessage(
        "メールアドレス形式で入力してください。"
      );
      return;
    }

    if (
      myEmail &&
      email === myEmail.toLowerCase()
    ) {
      setMessage(
        "自分自身は既に表示対象に含まれています。"
      );
      return;
    }

    if (
      displayUsers.some(
        (user) => user.email === email
      )
    ) {
      setMessage(
        "既に追加済みのユーザです。"
      );
      return;
    }

    setMessage(
      "ユーザ情報を取得中です..."
    );

    try {
      const accessToken =
        await getAccessToken();

      const profile =
        await getUserProfile(
          accessToken,
          email
        );

      const displayName =
        formatDisplayName(
          profile.displayName ||
            profile.mail ||
            profile.userPrincipalName ||
            email
        );

      const newUsers = [
        ...displayUsers,
        {
          email,
          displayName,
        },
      ];

      setDisplayUsers(newUsers);
saveSettings(newUsers, showSidebar, wrapText,showLocation);

      setEmailInput("");

      setMessage(
        "ユーザを追加しました。"
      );
    } catch (error) {
      console.error(error);

      setMessage(
        "ユーザ名を取得できませんでした。User.ReadBasic.All の権限追加、管理者同意、再ログインを確認してください。"
      );
    }
  };

  const removeUser = (
    email: string
  ) => {
    const newUsers =
      displayUsers.filter(
        (user) =>
          user.email !== email
      );

    setDisplayUsers(newUsers);
 saveSettings(
  newUsers,
  showSidebar,
  wrapText,
  showLocation
);

    setEventsByUser((current) => {
      const copy = { ...current };
      delete copy[email];
      return copy;
    });

    setMessage(
      "ユーザを削除しました。"
    );
  };

  const changeWeek = (
  newOffset: number
) => {
  setWeekOffset(newOffset);

  // スケジュール表をクリア
  setEventsByUser({});
};

  const moveUser = (
  index: number,
  direction: "up" | "down"
) => {
  const newUsers = [...displayUsers];

  const targetIndex =
    direction === "up"
      ? index - 1
      : index + 1;

  if (
    targetIndex < 0 ||
    targetIndex >= newUsers.length
  ) {
    return;
  }

  [
    newUsers[index],
    newUsers[targetIndex],
  ] = [
    newUsers[targetIndex],
    newUsers[index],
  ];

  setDisplayUsers(newUsers);
saveSettings(
  newUsers,
  showSidebar,
  wrapText,
  showLocation
);
};

  const dates = useMemo(() => {
    const today = new Date();

    const start = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );

    const dayOfWeek =
      start.getDay();

    start.setDate(
      start.getDate() - dayOfWeek
    );

    start.setDate(
      start.getDate() +
        weekOffset * 7
    );

    return Array.from(
      { length: 7 },
      (_, index) => {
        const date =
          new Date(start);

        date.setDate(
          start.getDate() + index
        );

        return date;
      }
    );
  }, [weekOffset]);

  const toDateKey = (
    date: Date
  ) => {
    const y = date.getFullYear();
    const m = String(
      date.getMonth() + 1
    ).padStart(2, "0");
    const d = String(
      date.getDate()
    ).padStart(2, "0");

    return `${y}-${m}-${d}`;
  };

  const todayKey =
    toDateKey(new Date());

  const getRangeStart = () => {
    const start =
      new Date(dates[0]);

    start.setHours(0, 0, 0, 0);

    return start;
  };

  const getRangeEnd = () => {
    const end =
      new Date(dates[6]);

    end.setDate(
      end.getDate() + 1
    );

    end.setHours(0, 0, 0, 0);

    return end;
  };

  const loadSchedule = async () => {
    if (!account) {
      return;
    }

    setLoadingSchedule(true);

    setMessage(
      "予定を取得中です..."
    );

    try {
      const accessToken =
        await getAccessToken();

      const rangeStart =
        getRangeStart();

      const rangeEnd =
        getRangeEnd();

      const nextEventsByUser: EventsByUser = {};

      const myEvents =
        await getMyEvents(
          accessToken,
          rangeStart,
          rangeEnd
        );

      nextEventsByUser[myEmail] =
        myEvents.map((event: any) => ({
          id: event.id,
          userEmail: myEmail,
          subject:
            event.subject ||
            "件名なし",
          location:
            event.location?.displayName ||
            "",
          start: event.start,
          end: event.end,
          showAs: event.showAs,
        }));

      for (const user of displayUsers) {
        try {
          const scheduleItems =
            await getScheduleForUser(
              accessToken,
              user.email,
              rangeStart,
              rangeEnd
            );

          const convertedEvents =
            scheduleItems.map(
              (
                item: any,
                index: number
              ) => ({
                id:
                  `${user.email}-${index}-${item.start?.dateTime}`,
                userEmail:
                  user.email,
                subject:
                  item.subject ||
                  item.status ||
                  "予定あり",
                location:
                  item.location ||
                  "",
                start: item.start,
                end: item.end,
                showAs:
                  item.status,
                status:
                  item.status,
              })
            );

          nextEventsByUser[user.email] =
            convertedEvents;
        } catch (error) {
          console.error(error);

          nextEventsByUser[user.email] = [
            {
              id: `error-${user.email}`,
              userEmail: user.email,
              subject:
                "予定を取得できませんでした。",
              location:
                "共有権限、Graph権限、または対象メールアドレスを確認してください。",
              start: {
                dateTime: "",
              },
              end: {
                dateTime: "",
              },
              isError: true,
            },
          ];
        }
      }

      setEventsByUser(nextEventsByUser);

      setMessage(
        "予定取得が完了しました。"
      );

      setLoadingSchedule(false);

    } catch (error) {
      console.error(error);

      setMessage(
        "予定取得でエラーが発生しました。"
      );

      setLoadingSchedule(false);

    }
  };

useEffect(() => {
  if (!account || !initialized) {
    return;
  }

  loadSchedule();
}, [
  account,
  initialized,
]);

  const usersForBoard = useMemo(() => {
    const users: DisplayUser[] = [];

    if (myEmail) {
      users.push({
        email: myEmail,
        displayName: myDisplayName,
      });
    }

    return [
      ...users,
      ...displayUsers,
    ];
  }, [
    myEmail,
    myDisplayName,
    displayUsers,
  ]);

  const formatDateHeader = (
    date: Date
  ) => {
    const dayNames = [
      "日",
      "月",
      "火",
      "水",
      "木",
      "金",
      "土",
    ];

    return `${
      date.getMonth() + 1
    }/${date.getDate()}(${
      dayNames[date.getDay()]
    })`;
  };
const dateKeyToLocalDate = (
  dateKey: string
) => {
  const [year, month, day] =
    dateKey.split("-").map(Number);

  return new Date(
    year,
    month - 1,
    day
  );
};

const isEventOnDate = (
  event: CalendarEvent,
  dateKey: string
) => {
  if (
    !event.start?.dateTime ||
    !event.end?.dateTime
  ) {
    return false;
  }

  const eventStart = new Date(
    event.start.dateTime
  );

  const eventEnd = new Date(
    event.end.dateTime
  );

  const dayStart =
    dateKeyToLocalDate(dateKey);

  const dayEnd =
    new Date(dayStart);

  dayEnd.setDate(
    dayEnd.getDate() + 1
  );

  return (
    eventStart < dayEnd &&
    eventEnd > dayStart
  );
};

const formatTime = (
  value?: string
) => {
  if (!value) {
    return "";
  }

  const time =
    value.split("T")[1] || "";

  return time.slice(0, 5);
};

const formatEventTime = (
  event: CalendarEvent
) => {
  const start = formatTime(
    event.start?.dateTime
  );

  const end = formatTime(
    event.end?.dateTime
  );

  if (
    start === "00:00" &&
    end === "00:00"
  ) {
    return "終日";
  }

  return `${start}-${end}`;
};

const getCellEvents = (
  userEmail: string,
  dateKey: string
) => {
  const events =
    eventsByUser[userEmail] || [];

  return events.filter((event) => {
    if (event.isError) {
      return (
        dateKey ===
        toDateKey(dates[0])
      );
    }

    return isEventOnDate(
      event,
      dateKey
    );
  });
};

  const getStatusLabel = (
    status?: string
  ) => {
    switch (status) {
      case "free":
        return "空き";
      case "tentative":
        return "仮予定";
      case "busy":
        return "予定あり";
      case "oof":
        return "外出中";
      case "workingElsewhere":
        return "他の場所で作業";
      default:
        return status || "-";
    }
  };

  const getStatusColor = (
    status?: string,
    isError?: boolean
  ) => {
    if (isError) {
      return "#dc2626";
    }

    switch (status) {
      case "free":
        return "#16a34a";
      case "tentative":
        return "#f59e0b";
      case "busy":
        return "#2563eb";
      case "oof":
        return "#dc2626";
      case "workingElsewhere":
        return "#7c3aed";
      default:
        return "#64748b";
    }
  };

  return (
    <div
      style={{
        padding: 10,
        fontFamily:
          "Segoe UI, Yu Gothic UI, Meiryo, sans-serif",
        background: "#f3f4f6",
        minHeight: "100vh",
        boxSizing: "border-box",
      }}
    >
<h1
  style={{
    margin: 0,
    padding: "7px 12px",
    marginBottom: 8,

    background: "#2563eb",
    color: "#ffffff",

    borderRadius: 6,

    position: "relative",

    textAlign: "center",

    fontSize: 18,
    fontWeight: 600,
  }}
>
  DIT Schedule Board

  <span
    style={{
      position: "absolute",

      right: 12,
      top: "50%",

      transform:
        "translateY(-50%)",

      fontSize: 16,

      color: "#bfdbfe",

      fontWeight: 400,

      whiteSpace: "nowrap",

      fontFamily:
        "Consolas, monospace",
    }}
  >
    BUILD_VERSION :
    {" "}
    {__BUILD_TIME__}
  </span>
</h1>

      {!account ? (
        <button
          onClick={signIn}
          style={primaryButtonStyle}
        >
          Microsoft 365 ログイン
        </button>
      ) : (
        <div
          style={{
            display: "flex",
            gap: 8,
            alignItems: "flex-start",
          }}
        >
          <div
            style={{
              width: showSidebar
                ? 240
                : 36,
              minWidth: showSidebar
                ? 240
                : 36,
              border:
                "1px solid #d1d5db",
              borderRadius: 6,
              background: "#f9fafb",
              padding: showSidebar
                ? 8
                : 4,
              boxSizing: "border-box",
              boxShadow:
                "0 1px 3px rgba(0,0,0,0.08)",
            }}
          >
            <button
              onClick={() => {
  const next =
    !showSidebar;

  setShowSidebar(next);

  saveSettings(displayUsers, next, wrapText,showLocation);
}
              }
              style={{
                ...buttonStyle,
                width: "100%",
                marginBottom: 8,
              }}
            >
              {showSidebar
                ? "◀隠す"
                : "▶開く"}
            </button>

            {showSidebar && (
              <>
                <div
                  style={{
                    marginBottom: 8,
                    paddingBottom: 8,
                    borderBottom:
                      "1px solid #ddd",
                    fontSize: 12,
                  }}
                >
                  <div
                    style={{
                      fontWeight: "bold",
                      marginBottom: 2,
                    }}
                  >
                    ログイン中
                  </div>

                  <div>
                    {myDisplayName}
                  </div>

                  <div
                    title={myEmail}
                    style={{
                      color: "#666",
                      overflow: "hidden",
                      whiteSpace: "nowrap",
                      textOverflow:
                        "ellipsis",
                    }}
                  >
                    {myEmail}
                  </div>
                </div>

                <h4
                  style={{
                    marginTop: 0,
                    marginBottom: 6,
                    fontSize: 13,
                  }}
                >
                  ユーザ追加
                </h4>

<div
  style={{
    position: "relative",
  }}
>
<input
  value={emailInput}
  onChange={(e) =>
    setEmailInput(
      e.target.value
    )
  }
    onFocus={() => {
      if (
        userCandidates.length > 0
      ) {
        setShowUserCandidates(true);
      }
    }}
    placeholder="メールアドレスまたは名前"
    style={{
      width: "100%",
      boxSizing:
        "border-box",
      padding: 5,
      border:
        "1px solid #d1d5db",
      borderRadius: 4,
      fontSize: 12,
    }}
  />

  {showUserCandidates && (
    <div
      style={{
        position: "absolute",
        top: 30,
        left: 0,
        right: 0,
        zIndex: 1000,

        background: "#ffffff",
        border:
          "1px solid #cbd5e1",
        borderRadius: 4,
        boxShadow:
          "0 4px 10px rgba(0,0,0,0.15)",

        maxHeight: 220,
        overflowY: "auto",
      }}
    >
      {isSearchingUsers ? (
        <div
          style={{
            padding: 8,
            fontSize: 12,
            color: "#64748b",
          }}
        >
          検索中...
        </div>
      ) : userCandidates.length ===
        0 ? (
        <div
          style={{
            padding: 8,
            fontSize: 12,
            color: "#64748b",
          }}
        >
          候補がありません
        </div>
      ) : (
        userCandidates.map(
          (candidate) => (
            <button
              key={
                candidate.email
              }
              type="button"
              onClick={() =>
                selectUserCandidate(
                  candidate
                )
              }
              style={{
                width: "100%",
                padding:
                  "6px 8px",
                border: "none",
                borderBottom:
                  "1px solid #e5e7eb",
                background:
                  "#ffffff",
                textAlign: "left",
                cursor:
                  "pointer",
              }}
            >
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 600,
                  color:
                    "#111827",
                  overflow:
                    "hidden",
                  whiteSpace:
                    "nowrap",
                  textOverflow:
                    "ellipsis",
                }}
              >
                {
                  candidate.displayName
                }
              </div>

              <div
                style={{
                  fontSize: 11,
                  color:
                    "#64748b",
                  overflow:
                    "hidden",
                  whiteSpace:
                    "nowrap",
                  textOverflow:
                    "ellipsis",
                }}
              >
                {
                  candidate.email
                }
              </div>
            </button>
          )
        )
      )}
    </div>
  )}
</div>

                <button
                  onClick={addUserByEmail}
                  style={{
                    ...buttonStyle,
                    width: "100%",
                    marginTop: 4,
                  }}
                >
                  追加
                </button>

                <p
                  style={{
                    color: "#555",
                    fontSize: 12,
                    lineHeight: 1.3,
                    minHeight: 16,
                    marginTop: 6,
                    marginBottom: 6,
                  }}
                >
                  {message}
                </p>

                <hr />

                <h4
                  style={{
                    marginTop: 8,
                    marginBottom: 6,
                    fontSize: 13,
                  }}
                >
                  表示対象ユーザ
                </h4>

                <ul
                  style={{
                    paddingLeft: 0,
                    margin: 0,
                    listStyle: "none",
                  }}
                >
                  {displayUsers.map(
                    (user, index) => (
                      <li
                        key={user.email}
                        style={{
                          marginBottom: 4,
                          fontSize: 12,
                          display: "flex",
                          alignItems: "center",
                          gap: 4,
                          overflow: "hidden",
                        }}
                      >
                        <button
                          onClick={() =>
                            removeUser(
                              user.email
                            )
                          }
                          style={{
                            ...buttonStyle,
                            fontSize: 10,
                            padding:
                              "1px 4px",
                            flexShrink: 0,
                          }}
                        >
                          削除
                        </button>
<button
  onClick={() =>
    moveUser(index, "up")
  }
  style={{
    ...buttonStyle,
    fontSize: 10,
    padding: "1px 4px",
  }}
  disabled={index === 0}
>
  ↑
</button>

<button
  onClick={() =>
    moveUser(index, "down")
  }
  style={{
    ...buttonStyle,
    fontSize: 10,
    padding: "1px 4px",
  }}
  disabled={
    index ===
    displayUsers.length - 1
  }
>
  ↓
</button>
                        <span
                          style={{
                            fontWeight:
                              "bold",
                            flexShrink: 0,
                          }}
                        >
                          {formatDisplayName(
                            user.displayName
                          )}
                        </span>

                        <span
                          title={user.email}
                          style={{
                            color: "#666",
                            overflow:
                              "hidden",
                            whiteSpace:
                              "nowrap",
                            textOverflow:
                              "ellipsis",
                            minWidth: 0,
                            flex: 1,
                          }}
                        >
                          {user.email}
                        </span>
                      </li>
                    )
                  )}
                </ul>
              </>
            )}
          </div>

          <div
            style={{
              flex: 1,
              minWidth: 0,
            }}
          >
            <div
              style={{
                marginBottom: 8,
                display: "flex",
                alignItems: "center",
                gap: 4,
                flexWrap: "wrap",
                background: "#ffffff",
                border:
                  "1px solid #d1d5db",
                borderRadius: 6,
                padding: 6,
                boxShadow:
                  "0 1px 2px rgba(0,0,0,0.06)",
              }}
            >

<div
  style={{
    display: "flex",
    alignItems: "center",
    gap: 10,

    padding: "3px 8px",

    border: "1px solid #cbd5e1",
    borderRadius: 4,

    background: "#ffffff",

    fontSize: 11,

    marginRight: 12,

    whiteSpace: "nowrap",
  }}
>

  <span
    style={{
      display: "flex",
      alignItems: "center",
      gap: 3,
    }}
  >
    <span
      style={{
        width: 8,
        height: 8,
        borderRadius: "50%",
        background: "#2563eb",
      }}
    />
    予定
  </span>

  <span
    style={{
      display: "flex",
      alignItems: "center",
      gap: 3,
    }}
  >
    <span
      style={{
        width: 8,
        height: 8,
        borderRadius: "50%",
        background: "#f59e0b",
      }}
    />
    仮
  </span>

  <span
    style={{
      display: "flex",
      alignItems: "center",
      gap: 3,
    }}
  >
    <span
      style={{
        width: 8,
        height: 8,
        borderRadius: "50%",
        background: "#dc2626",
      }}
    />
    外出
  </span>

  <span
    style={{
      display: "flex",
      alignItems: "center",
      gap: 3,
    }}
  >
    <span
      style={{
        width: 8,
        height: 8,
        borderRadius: "50%",
        background: "#16a34a",
      }}
    />
    空き
  </span>
</div>
              <span
                style={{
                  fontWeight: "bold",
                  marginRight: 12,
                  fontSize: 13,
                  color: "#374151",
                }}
              >
                {formatDateHeader(dates[0])}
                ～
                {formatDateHeader(dates[6])}
              </span>

<button
  onClick={() =>
    changeWeek(
      weekOffset - 1
    )
  }
  style={buttonStyle}
>
  ← 前週
</button>

<button
  onClick={() =>
    changeWeek(0)
  }
  style={buttonStyle}
>
  今週
</button>

<button
  onClick={() =>
    changeWeek(
      weekOffset + 1
    )
  }
  style={buttonStyle}
>
  翌週 →
</button>

              <button
                onClick={loadSchedule}
                style={primaryButtonStyle}
                disabled={loadingSchedule}
              >
                {loadingSchedule ? "取得中..." : "予定取得"}
              </button>

              <label
                style={{
                  marginLeft: 8,
                  fontSize: 12,
                  display: "flex",
                  alignItems: "center",
                  gap: 4,
                  color: "#374151",
                }}
              >
  <input
    type="checkbox"
    checked={wrapText}
    onChange={(e) => {
      const next =
        e.target.checked;

      setWrapText(next);

      saveSettings(
        displayUsers,
        showSidebar,
        next,
        showLocation
      );
    }}
  />
  折り返し表示
</label>
<label
  style={{
    marginLeft: 8,
    fontSize: 12,
    display: "flex",
    alignItems: "center",
    gap: 4,
    color: "#374151",
  }}
>
  <input
    type="checkbox"
    checked={showLocation}
    onChange={(e) => {
      const next =
        e.target.checked;

      setShowLocation(next);

      saveSettings(
        displayUsers,
        showSidebar,
        wrapText,
        next
      );
    }}
  />
  場所表示
</label>
<button
  onClick={signOut}
  style={{
    ...buttonStyle,
    marginLeft: "auto",
  }}
>
  ログアウト
</button>
            </div>
{loadingSchedule ? (
  <div
    style={{
      background: "#ffffff",
      border: "1px solid #d1d5db",
      borderRadius: 6,
      padding: 20,
      textAlign: "center",
      color: "#374151",
    }}
  >
    予定を取得中です...
  </div>
) : (
  <div
    style={{
      overflowX: "auto",
      overflowY: "auto",
      height: "calc(100vh - 140px)",
      background: "#ffffff",
      border: "1px solid #d1d5db",
      borderRadius: 6,
      boxShadow:
        "0 1px 3px rgba(0,0,0,0.08)",
    }}
  >
  <table
                  style={{
                    borderCollapse:
                      "collapse",
                    width: "100%",
                    tableLayout: "fixed",
                  }}
                >
                  <thead>
                    <tr>
                      <th
                        style={{
                          ...thStyle,
                          width: 95,
                          minWidth: 95,
                          maxWidth: 95,
                          left: 0,
                          top: 0,
                          zIndex: 30,
                          position: "sticky",
                          background:
                            "#e5e7eb",
                        }}
                      >
                        ユーザ
                      </th>

                      {dates.map((date) => (
<th
  key={toDateKey(date)}
  style={{
    ...thStyle,
    top: 0,
    zIndex: 20,

    background:
      toDateKey(date) ===
      todayKey
        ? "#dbeafe"
        : "#e5e7eb",

    color:
      date.getDay() === 0
        ? "#dc2626" // 日曜=赤
        : date.getDay() === 6
        ? "#2563eb" // 土曜=青
        : "#111827",
  }}
>
  {formatDateHeader(date)}
</th>
                      ))}
                    </tr>
                  </thead>

                  <tbody>
                    {usersForBoard.map(
                      (
                        user,
                        rowIndex
                      ) => {
                        const rowBackground =
                          rowIndex % 2 ===
                          0
                            ? "#f8fcf8"
                            : "#f8faff";

                        return (
                          <tr key={user.email}>
                            <td
                              title={user.email}
                              style={{
                                ...tdStyle,
                                fontWeight:
                                  "bold",
                                width: 95,
                                minWidth: 95,
                                maxWidth: 95,
                                position:
                                  "sticky",
                                left: 0,
                                zIndex: 5,
                                background:
                                  rowBackground,
whiteSpace: "normal",
wordBreak: "break-word",
overflow: "visible",
lineHeight: "1.2",
                              }}
                            >
                              {formatDisplayName(
                                user.displayName
                              )}
                            </td>

                            {dates.map((date) => {
                              const dateKey =
                                toDateKey(date);

                              const cellEvents =
                                getCellEvents(
                                  user.email,
                                  dateKey
                                );

                              return (
                                <td
                                  key={
                                    user.email +
                                    dateKey
                                  }
                                  style={{
                                    ...tdStyle,
                                    background:
                                      dateKey ===
                                      todayKey
                                        ? rowIndex %
                                            2 ===
                                          0
                                          ? "#dff1df"
                                          : "#dbe8ff"
                                        : rowBackground,
                                  }}
                                >
                                  {cellEvents.length ===
                                  0 ? (
                                    <div
                                      style={{
                                        textAlign:
                                          "center",
                                        color:
                                          "#94a3b8",
                                        fontSize: 12,
                                        lineHeight:
                                          "1.2",
                                      }}
                                    >
                                      －
                                    </div>
                                  ) : (
                                    cellEvents.map(
                                      (event) => {
                                        const fullTitle =
                                          event.subject ||
                                          "件名なし";

                                        const fullLocation =
                                          event.location ||
                                          "";

                                        const tooltip =
                                          [
                                            fullTitle,
                                            fullLocation,
                                          ]
                                            .filter(
                                              Boolean
                                            )
                                            .join(
                                              "\n"
                                            );

                                        return (
                                        <div
                                          key={event.id}
                                          title={tooltip}
                                          style={{
                                            display: "flex",
                                            alignItems: "flex-start",
                                            gap: 4,

background: "#ffffff",

borderRadius: 3,

padding: 3,

marginBottom: 3,
                                            border: "1px solid #99acc0",

                                            fontSize: 11,
                                            lineHeight: "1.25",
                                          }}
                                        >
                                            <span
                                              title={getStatusLabel(
                                                event.showAs
                                              )}
                                              style={{
                                                width: 7,
                                                height: 7,
                                                borderRadius:
                                                  "50%",
                                                flexShrink: 0,
                                                marginTop: 4,
                                                backgroundColor:
                                                  getStatusColor(
                                                    event.showAs,
                                                    event.isError
                                                  ),
                                              }}
                                            />

                                            <div
                                              style={{
                                                minWidth: 0,
                                                flex: 1,
                                              }}
                                            >
                                              {!event.isError && (
                                                <div
                                                  style={{
                                                    color:
                                                      "#334155",
                                                    fontSize: 10,
                                                    fontWeight: 600,
                                                    whiteSpace:
                                                      "nowrap",
                                                    overflow:
                                                      "hidden",
                                                    textOverflow:
                                                      "ellipsis",
                                                    lineHeight:
                                                      "1.15",
                                                  }}
                                                >
                                                  {formatEventTime(event)}
                                                </div>
                                              )}

                                              <div
                                                style={{
                                                  overflow:
                                                    wrapText
                                                      ? "visible"
                                                      : "hidden",
                                                  whiteSpace:
                                                    wrapText
                                                      ? "normal"
                                                      : "nowrap",
                                                  wordBreak:
                                                    wrapText
                                                      ? "break-word"
                                                      : "normal",
                                                  textOverflow:
                                                    wrapText
                                                      ? "clip"
                                                      : "ellipsis",
                                                  fontWeight: 600,
                                                  color:
                                                    event.isError
                                                      ? "#b91c1c"
                                                      : "#111827",
                                                  lineHeight:
                                                    "1.25",
                                                }}
                                              >
                                                {fullTitle}
                                              </div>

                                              {showLocation && fullLocation && (
                                                <div
                                                  style={{
                                                    overflow:
                                                      wrapText
                                                        ? "visible"
                                                        : "hidden",
                                                    whiteSpace:
                                                      wrapText
                                                        ? "normal"
                                                        : "nowrap",
                                                    wordBreak:
                                                      wrapText
                                                        ? "break-word"
                                                        : "normal",
                                                    textOverflow:
                                                      wrapText
                                                        ? "clip"
                                                        : "ellipsis",
                                                    color:
                                                      "#64748b",
                                                    lineHeight:
                                                      "1.2",
                                                  }}
                                                >
                                                  {fullLocation}
                                                </div>
                                              )}
                                            </div>
                                          </div>
                                        );
                                      }
                                    )
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      }
                    )}
                  </tbody>
                </table>
              </div>
          )}
          </div>
        </div>
      )}
    </div>
  );
}

const thStyle: CSSProperties = {
  border: "1px solid #cbd5e1",
  padding: 5,
  textAlign: "center",
  verticalAlign: "middle",
  position: "sticky",
  background: "#e5e7eb",
  color: "#111827",
  fontWeight: 600,
  fontSize: 12,
  lineHeight: "1.2",
};

const tdStyle: CSSProperties = {
  border: "1px solid #d1d5db",
  padding: 3,
  verticalAlign: "top",
  width: "calc((100% - 95px) / 7)",
};

export default App;