// calendar_id is only useful if it reaches Graph on every event path. These tests assert the
// URLs the client sends, so a tool that drops the ID — and silently writes to the default
// calendar instead — fails here rather than on someone's real calendar.

import type { AuthStrategy } from "@sapientsai/ms-graph-core"
import { Right } from "functype/either"
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"

import { initializeGraphClient } from "../src/client/graph-client"
import {
  createEvent,
  deleteEvent,
  getEvent,
  listCalendars,
  listCalendarView,
  listEvents,
  updateEvent,
} from "../src/tools/calendar-tools"

const CAL = "AAMkCal-1="
const event = { id: "evt-1", subject: "Swimming carnival", start: { dateTime: "2026-11-06T09:00:00" } }

const fetchSpy = vi.fn()
const sentUrl = (call = 0) => new URL(fetchSpy.mock.calls[call]![0] as string)
const sentMethod = (call = 0) => (fetchSpy.mock.calls[call]![1] as RequestInit | undefined)?.method ?? "GET"

beforeEach(() => {
  fetchSpy.mockReset()
  fetchSpy.mockImplementation(() =>
    Promise.resolve(new Response(JSON.stringify({ ...event, value: [event] }), { status: 200 })),
  )
  vi.stubGlobal("fetch", fetchSpy)
  const auth: AuthStrategy = { getAccessToken: () => Promise.resolve(Right("T")) }
  initializeGraphClient(auth)
})

afterEach(() => {
  vi.unstubAllGlobals()
})

describe("calendar selection", () => {
  it("list_calendars reads /me/calendars and reports each calendar's ID, owner and edit rights", async () => {
    fetchSpy.mockResolvedValueOnce(
      new Response(
        JSON.stringify({
          value: [
            { id: "cal-default", name: "Calendar", isDefaultCalendar: true, canEdit: true },
            {
              id: CAL,
              name: "School Calendar",
              isDefaultCalendar: false,
              canEdit: true,
              owner: { name: "Bel", address: "bel@example.com" },
            },
            { id: "cal-ro", name: "Holidays", canEdit: false },
          ],
        }),
      ),
    )

    const result = await listCalendars()

    expect(sentUrl().pathname).toMatch(/\/me\/calendars$/)
    expect(sentUrl().searchParams.get("$select")).toBe("id,name,isDefaultCalendar,canEdit,owner")
    expect(result.value).toContain("**Calendar** [default] (ID: cal-default)")
    expect(result.value).toContain(`**School Calendar** — owner bel@example.com (ID: ${CAL})`)
    expect(result.value).toContain("**Holidays** [read-only] (ID: cal-ro)")
  })

  describe("without calendar_id, events stay on the default calendar", () => {
    it.each([
      [
        "create_event",
        () => createEvent({ subject: "s", start: "2026-11-06T09:00", end: "2026-11-06T10:00" }),
        "POST",
        "/me/events",
      ],
      ["list_events", () => listEvents({}), "GET", "/me/events"],
      ["list_events (all pages)", () => listEvents({ fetch_all_pages: true }), "GET", "/me/events"],
      ["get_event", () => getEvent({ event_id: "evt-1" }), "GET", "/me/events/evt-1"],
      ["update_event", () => updateEvent({ event_id: "evt-1", subject: "s" }), "PATCH", "/me/events/evt-1"],
      ["delete_event", () => deleteEvent({ event_id: "evt-1" }), "DELETE", "/me/events/evt-1"],
      [
        "list_calendar_view",
        () => listCalendarView({ start_date_time: "2026-11-01T00:00:00Z", end_date_time: "2026-11-08T00:00:00Z" }),
        "GET",
        "/me/calendarView",
      ],
    ] as const)("%s", async (_name, run, method, path) => {
      const result = await run()

      expect(result.isRight()).toBe(true)
      expect(sentMethod()).toBe(method)
      expect(sentUrl().pathname).toMatch(new RegExp(`/v1\\.0${path}$`))
    })
  })

  describe("with calendar_id, events go to that calendar", () => {
    const calPath = `/me/calendars/${CAL}`

    it.each([
      [
        "create_event",
        () => createEvent({ subject: "s", start: "2026-11-06T09:00", end: "2026-11-06T10:00", calendar_id: CAL }),
        "POST",
        `${calPath}/events`,
      ],
      ["list_events", () => listEvents({ calendar_id: CAL }), "GET", `${calPath}/events`],
      [
        "list_events (all pages)",
        () => listEvents({ fetch_all_pages: true, calendar_id: CAL }),
        "GET",
        `${calPath}/events`,
      ],
      ["get_event", () => getEvent({ event_id: "evt-1", calendar_id: CAL }), "GET", `${calPath}/events/evt-1`],
      [
        "update_event",
        () => updateEvent({ event_id: "evt-1", subject: "s", calendar_id: CAL }),
        "PATCH",
        `${calPath}/events/evt-1`,
      ],
      ["delete_event", () => deleteEvent({ event_id: "evt-1", calendar_id: CAL }), "DELETE", `${calPath}/events/evt-1`],
      [
        "list_calendar_view",
        () =>
          listCalendarView({
            start_date_time: "2026-11-01T00:00:00Z",
            end_date_time: "2026-11-08T00:00:00Z",
            calendar_id: CAL,
          }),
        "GET",
        `${calPath}/calendarView`,
      ],
    ] as const)("%s", async (_name, run, method, path) => {
      const result = await run()

      expect(result.isRight()).toBe(true)
      expect(sentMethod()).toBe(method)
      expect(decodeURIComponent(sentUrl().pathname)).toMatch(
        new RegExp(`/v1\\.0${path.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}$`),
      )
    })

    it("list_calendar_view keeps its window on the calendar-scoped path", async () => {
      await listCalendarView({
        start_date_time: "2026-11-01T00:00:00Z",
        end_date_time: "2026-11-08T00:00:00Z",
        calendar_id: CAL,
      })

      expect(sentUrl().searchParams.get("startDateTime")).toBe("2026-11-01T00:00:00Z")
      expect(sentUrl().searchParams.get("endDateTime")).toBe("2026-11-08T00:00:00Z")
    })
  })
})
