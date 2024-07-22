"use client";

import Image from "next/image";
import "./auth";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import { ensureScope, getToken, signIn } from "./auth";
import { useState } from "react";

export default function Home() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [userName, setUserName] = useState("");
  const [events, setEvents] = useState([]);
  async function displayUI() {
    await signIn();
    // Display info from user profile
    await getUser().then((user) => {
      setUserName(user && user!.displayName);
    });
  }

  async function displayEvents() {
    var eventsContainer = await getEvents();

    setEvents(eventsContainer!.value);
  }

  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      {!userName && (
        <a id="signin" href="#" onClick={displayUI}>
          <img src="/microsoft_logo.svg" alt="Sign in with Microsoft" />
          Sign in with Microsoft
        </a>
      )}

      {userName && (
        <>
          <div id="content" className="none;">
            <h4>
              Welcome <span id="userName"> {userName} </span>
            </h4>
          </div>
          {!events && (
            <button id="btnShowEvents" onClick={displayEvents}>
              Show events
            </button>
          )}

          <div id="eventWrapper">
            {!events && <p>Failed to retrieve from Microsoft Graph</p>}
            {events && events.length > 0 && (
              <>
                <p>
                  Your events retrieved from Microsoft Graph for the day week:
                </p>
                <ul>
                  {events.map((idx: any, elm: any) => (
                    <li id="events" key={elm.subject}>{`${
                      elm.subject
                    } - From  ${new Date(
                      elm.start.dateTime
                    ).toLocaleString()} to ${new Date(
                      elm.end.dateTime
                    ).toLocaleString()} `}</li>
                  ))}
                </ul>
              </>
            )}
            {(!events || events.length < 1) && (
              <p>No events retrieved from Microsoft Graph</p>
            )}
          </div>
        </>
      )}
    </main>
  );
}

// Create an authentication provider
const authProvider = {
  getAccessToken: async () => {
    // Call getToken in auth.js
    return await getToken();
  },
};

// Initialize the Graph client
const graphClient: Client = Client.initWithMiddleware({ authProvider });

//Get user info from Graph
async function getUser() {
  ensureScope("user.read");
  return await graphClient.api("/me").select("id,displayName").get();
}

async function getEvents(): Promise<any> {
  ensureScope("Calendars.read");
  const dateNow = new Date();
  const dateNextWeek = new Date();
  dateNextWeek.setDate(dateNextWeek.getDate() + 1);
  const query = `startDateTime=${dateNow.toISOString()}&endDateTime=${dateNextWeek.toISOString()}`;

  return await graphClient
    .api("/me/calendarView")
    .query(query)
    .select("subject,start,end")
    .orderby(`start/DateTime`)
    .get();
}
