import { Some } from "functype"
import { Left, Right } from "functype/either"
import { beforeEach, describe, expect, it, vi } from "vitest"

vi.mock("../src/client/graph-client", () => ({
  getGraphClient: vi.fn(),
}))

import { getGraphClient } from "../src/client/graph-client"
import { createTodoList, createTodoTask, deleteTodoTask, updateTodoTask } from "../src/tools/todo-tools"

const mockClient = {
  createTodoTask: vi.fn(),
  updateTodoTask: vi.fn(),
  getTodoTask: vi.fn(),
  deleteTodoTask: vi.fn(),
  listTodoLists: vi.fn(),
  createTodoList: vi.fn(),
  requestPaginated: vi.fn(),
}

const QUARTERLY = {
  pattern: { type: "absoluteMonthly", interval: 3, dayOfMonth: 1 },
  range: { type: "noEnd", startDate: "2026-10-01" },
}

const LIST_ID = "list-1"

beforeEach(() => {
  vi.clearAllMocks()
  vi.mocked(getGraphClient).mockReturnValue(Some(mockClient as never))
})

describe("todo-tools", () => {
  describe("createTodoList", () => {
    it("should create a list when the name is free", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "l1", displayName: "Household" }]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      const result = await createTodoList({ display_name: "Renovation" })

      expect(result.isRight()).toBe(true)
      expect(mockClient.createTodoList).toHaveBeenCalledWith("Renovation")
    })

    // Graph happily creates a second list with the same displayName, and no client can
    // then tell the two apart — so tasks scatter across both.
    it("should refuse a name an existing list already uses, case-insensitively", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([{ id: "l1", displayName: "Household" }]))

      const result = await createTodoList({ display_name: "household" })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should check every page of lists, not just the first", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      await createTodoList({ display_name: "Renovation" })

      expect(mockClient.requestPaginated).toHaveBeenCalledWith("/me/todo/lists")
      expect(mockClient.listTodoLists).not.toHaveBeenCalled()
    })

    // A read failure leaves the duplicate check blind, so creating anyway could still
    // produce the duplicate the guard exists to prevent.
    it("should refuse to create when the existing lists cannot be read", async () => {
      mockClient.requestPaginated.mockResolvedValue(Left({ message: "Service unavailable" }))

      const result = await createTodoList({ display_name: "Renovation" })

      expect(result.isLeft()).toBe(true)
      expect(
        result.fold(
          (error) => error.message,
          () => "",
        ),
      ).toContain("Service unavailable")
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should refuse a blank name without calling Graph", async () => {
      const result = await createTodoList({ display_name: "   " })

      expect(result.isLeft()).toBe(true)
      expect(mockClient.requestPaginated).not.toHaveBeenCalled()
      expect(mockClient.createTodoList).not.toHaveBeenCalled()
    })

    it("should trim the name before sending it", async () => {
      mockClient.requestPaginated.mockResolvedValue(Right([]))
      mockClient.createTodoList.mockResolvedValue(Right({ id: "l2", displayName: "Renovation" }))

      await createTodoList({ display_name: "  Renovation  " })

      expect(mockClient.createTodoList).toHaveBeenCalledWith("Renovation")
    })
  })

  describe("createTodoTask", () => {
    it("should create a task with just a title", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      const result = await createTodoTask({ list_id: LIST_ID, title: "Buy milk" })
      expect(result.isRight()).toBe(true)
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, { title: "Buy milk" })
    })

    // Graph's todoTask.body is an itemBody whose bodyType enum is lowercase ("text" |
    // "html"). Sending "Text" is rejected with "Requested value 'Text' was not found",
    // which fails the whole create — not just the body.
    it("should send a lowercase bodyType so Graph accepts the body", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Call plumber" }))
      await createTodoTask({ list_id: LIST_ID, title: "Call plumber", body: "0400 000 000" })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Call plumber",
        body: { contentType: "text", content: "0400 000 000" },
      })
    })

    it("should pass due date and importance through", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Renew rego" }))
      await createTodoTask({
        list_id: LIST_ID,
        title: "Renew rego",
        due_date: "2026-10-01T00:00:00",
        importance: "high",
      })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Renew rego",
        dueDateTime: { dateTime: "2026-10-01T00:00:00", timeZone: "UTC" },
        importance: "high",
      })
    })

    it("should attach a recurrence and default its range to the due date", async () => {
      mockClient.createTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters" }))
      await createTodoTask({
        list_id: LIST_ID,
        title: "Clear the gutters",
        due_date: "2026-10-01T00:00:00",
        recurrence: { pattern: "absoluteMonthly", interval: 3, day_of_month: 1 },
      })
      expect(mockClient.createTodoTask).toHaveBeenCalledWith(LIST_ID, {
        title: "Clear the gutters",
        dueDateTime: { dateTime: "2026-10-01T00:00:00", timeZone: "UTC" },
        recurrence: {
          pattern: { type: "absoluteMonthly", interval: 3, dayOfMonth: 1 },
          range: { type: "noEnd", startDate: "2026-10-01" },
        },
      })
    })

    // To Do rolls a recurring task forward from its due date, so without one the
    // task repeats but never surfaces in Today.
    it("should refuse a recurring task with no due date", async () => {
      const result = await createTodoTask({
        list_id: LIST_ID,
        title: "Clear the gutters",
        recurrence: { pattern: "daily" },
      })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("due_date")
      expect(mockClient.createTodoTask).not.toHaveBeenCalled()
    })

    it("should not call Graph when the recurrence is incomplete", async () => {
      const result = await createTodoTask({
        list_id: LIST_ID,
        title: "Bin night",
        due_date: "2026-10-01T00:00:00",
        recurrence: { pattern: "weekly" },
      })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("days_of_week")
      expect(mockClient.createTodoTask).not.toHaveBeenCalled()
    })

    it("should surface a create failure as a UserError", async () => {
      mockClient.createTodoTask.mockResolvedValue(
        (await import("functype/either")).Left({ message: "Requested value 'Text' was not found." }),
      )
      const result = await createTodoTask({ list_id: LIST_ID, title: "Nope", body: "x" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to create task")
    })
  })

  describe("updateTodoTask", () => {
    it("should send a lowercase bodyType on update too", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", title: "Call plumber" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", body: "new notes" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        body: { contentType: "text", content: "new notes" },
      })
    })

    it("should only send the fields that were provided", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", status: "completed" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", { status: "completed" })
    })

    it("should add a recurrence to an existing task", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1" }))
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        recurrence: { pattern: "absoluteYearly", day_of_month: 1, month: 10, start_date: "2026-10-01" },
      })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        recurrence: {
          pattern: { type: "absoluteYearly", interval: 1, dayOfMonth: 1, month: 10 },
          range: { type: "noEnd", startDate: "2026-10-01" },
        },
      })
    })

    // Graph only clears a recurrence on an explicit null; omitting it leaves the
    // existing pattern untouched.
    it("should clear a recurrence with an explicit null", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1" }))
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", clear_recurrence: true })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", { recurrence: null })
    })

    // Graph rolls a recurring task forward on completion: the same id comes back as
    // notStarted with the next due date, and the completed occurrence becomes a
    // separate task. "Task updated" over a notStarted task reads as a failure.
    it("should explain the roll-forward when completing a recurring task", async () => {
      mockClient.updateTodoTask.mockResolvedValue(
        Right({
          id: "t1",
          title: "Clear the gutters",
          status: "notStarted",
          dueDateTime: { dateTime: "2027-01-01T00:00:00.0000000" },
          recurrence: QUARTERLY,
        }),
      )
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(result.isRight()).toBe(true)
      const text = result.value as string
      expect(text).toContain("Occurrence completed")
      expect(text).toContain("2027-01-01")
    })

    it("should say plainly when a non-recurring task is completed", async () => {
      mockClient.updateTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk", status: "completed" }))
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", status: "completed" })
      expect(result.value as string).toContain("Task completed.")
    })

    it("should reject setting and clearing a recurrence at once", async () => {
      const result = await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        clear_recurrence: true,
        recurrence: { pattern: "daily", start_date: "2026-10-01" },
      })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })
  })

  // A fake To Do list that behaves as Graph did live on 2026-10-28, for daily tasks:
  // a dueDateTime PATCH on a repeating task — with or without a recurrence beside
  // it — rolls the series forward a day and creates a second task on the requested
  // date; a recurrence PATCH moves the due date to the range's startDate.
  const fakeTodoList = (initial: ReadonlyArray<Record<string, unknown>>) => {
    const tasks = new Map(initial.map((t) => [t.id as string, { ...t }]))
    const dueOf = (date: string) => ({ dateTime: `${date}T00:00:00.0000000`, timeZone: "UTC" })
    const nextDay = (date: string) => new Date(Date.parse(`${date}T00:00:00Z`) + 864e5).toISOString().slice(0, 10)
    mockClient.getTodoTask.mockImplementation(async (_list: string, id: string) => Right({ ...tasks.get(id) }))
    mockClient.updateTodoTask.mockImplementation(async (_list: string, id: string, patch: Record<string, unknown>) => {
      const task = tasks.get(id)!
      const due = patch.dueDateTime as { dateTime: string } | undefined
      if (due && task.recurrence && patch.recurrence !== null) {
        tasks.set(`${id}-copy`, { ...task, id: `${id}-copy`, dueDateTime: dueOf(due.dateTime.slice(0, 10)) })
        const current = (task.dueDateTime as { dateTime: string }).dateTime.slice(0, 10)
        Object.assign(task, { dueDateTime: dueOf(nextDay(current)) })
        return Right({ ...task })
      }
      Object.assign(task, patch)
      if (due) task.dueDateTime = dueOf(due.dateTime.slice(0, 10))
      const recurrence = patch.recurrence as { range: { startDate: string } } | null | undefined
      if (recurrence) task.dueDateTime = dueOf(recurrence.range.startDate)
      return Right({ ...task })
    })
    return tasks
  }

  const DAILY = { pattern: { type: "daily", interval: 1 }, range: { type: "noEnd", startDate: "2026-10-30" } }
  const dueDates = (tasks: Map<string, Record<string, unknown>>) =>
    [...tasks.values()].map((t) => (t.dueDateTime as { dateTime: string }).dateTime.slice(0, 10))

  describe("updateTodoTask rescheduling a repeating task", () => {
    const repeating = () => [
      { id: "t1", title: "Water plants", dueDateTime: { dateTime: "2026-10-30T00:00:00.0000000" }, recurrence: DAILY },
    ]

    it("moves it with due_date alone, without creating a second task", async () => {
      const tasks = fakeTodoList(repeating())
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06T12:00:00Z" })
      expect(result.isRight()).toBe(true)
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.updateTodoTask.mock.calls[0]![2]).not.toHaveProperty("dueDateTime")
    })

    it("moves it with due_date and recurrence together, without creating a second task", async () => {
      const tasks = fakeTodoList(repeating())
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        recurrence: { pattern: "daily" },
      })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.getTodoTask).not.toHaveBeenCalled()
    })

    it("keeps the rest of the existing pattern when re-anchoring it", async () => {
      fakeTodoList([{ ...repeating()[0], recurrence: QUARTERLY }])
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2027-01-01" })
      expect(mockClient.updateTodoTask).toHaveBeenCalledWith(LIST_ID, "t1", {
        recurrence: { pattern: QUARTERLY.pattern, range: { type: "noEnd", startDate: "2027-01-01" } },
      })
    })

    it("still sends dueDateTime when the recurrence is being cleared", async () => {
      const tasks = fakeTodoList(repeating())
      await updateTodoTask({
        list_id: LIST_ID,
        task_id: "t1",
        due_date: "2026-11-06T12:00:00Z",
        clear_recurrence: true,
      })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.getTodoTask).not.toHaveBeenCalled()
    })

    it("sends dueDateTime for a one-off task", async () => {
      const tasks = fakeTodoList([
        { id: "t1", title: "Buy milk", dueDateTime: { dateTime: "2026-10-30T00:00:00.0000000" } },
      ])
      await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06T12:00:00Z" })
      expect(dueDates(tasks)).toEqual(["2026-11-06"])
      expect(mockClient.updateTodoTask.mock.calls[0]![2]).toHaveProperty("dueDateTime")
    })

    // Graph puts the task on the first date the pattern allows from startDate.
    it("says where the task landed when the pattern moved it off the asked date", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", recurrence: QUARTERLY }))
      mockClient.updateTodoTask.mockResolvedValue(
        Right({ id: "t1", dueDateTime: { dateTime: "2027-01-01T00:00:00.0000000" }, recurrence: QUARTERLY }),
      )
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-12-15" })
      expect(result.value as string).toContain("the next occurrence is 2027-01-01")
    })

    it("does not update when the task cannot be read first", async () => {
      mockClient.getTodoTask.mockResolvedValue(Left({ message: "ErrorItemNotFound" }))
      const result = await updateTodoTask({ list_id: LIST_ID, task_id: "t1", due_date: "2026-11-06" })
      expect(result.isLeft()).toBe(true)
      expect(mockClient.updateTodoTask).not.toHaveBeenCalled()
    })
  })

  describe("deleteTodoTask", () => {
    it("should delete a one-off task and name it", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      mockClient.deleteTodoTask.mockResolvedValue(Right({}))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isRight()).toBe(true)
      expect(result.value as string).toContain('Deleted "Buy milk"')
      expect(mockClient.deleteTodoTask).toHaveBeenCalledWith(LIST_ID, "t1")
    })

    // A task id addresses the whole series, not one occurrence, and Graph has no
    // restore — so this refusal is the only thing standing between a routine
    // tidy-up and silently destroying a recurring chore.
    it("should refuse a recurring task without force", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters", recurrence: QUARTERLY }))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isLeft()).toBe(true)
      const message = (result.value as Error).message
      expect(message).toContain("ends the whole series")
      expect(message).toContain("every 3 months")
      expect(message).toContain("clear_recurrence")
      expect(mockClient.deleteTodoTask).not.toHaveBeenCalled()
    })

    it("should delete a recurring task when forced, and say the series ended", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Clear the gutters", recurrence: QUARTERLY }))
      mockClient.deleteTodoTask.mockResolvedValue(Right({}))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1", force: true })
      expect(result.isRight()).toBe(true)
      expect(result.value as string).toContain("whole recurring series")
      expect(mockClient.deleteTodoTask).toHaveBeenCalledWith(LIST_ID, "t1")
    })

    it("should not delete when the task cannot be read first", async () => {
      mockClient.getTodoTask.mockResolvedValue((await import("functype/either")).Left({ message: "ErrorItemNotFound" }))
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "gone" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("before deleting")
      expect(mockClient.deleteTodoTask).not.toHaveBeenCalled()
    })

    it("should surface a delete failure as a UserError", async () => {
      mockClient.getTodoTask.mockResolvedValue(Right({ id: "t1", title: "Buy milk" }))
      mockClient.deleteTodoTask.mockResolvedValue(
        (await import("functype/either")).Left({ message: "ErrorAccessDenied" }),
      )
      const result = await deleteTodoTask({ list_id: LIST_ID, task_id: "t1" })
      expect(result.isLeft()).toBe(true)
      expect((result.value as Error).message).toContain("Failed to delete task")
    })
  })
})
