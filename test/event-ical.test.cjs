"use strict";

const assert = require("node:assert/strict");
const { pathToFileURL } = require("node:url");
const { test } = require("node:test");

const modulePath = process.env.TBSYNC_TEST_ICAL_MODULE;
const options = { skip: !modulePath && "Set TBSYNC_TEST_ICAL_MODULE to Thunderbird's Ical.sys.mjs." };

async function fixture() {
  const [{ default: ICAL }, { createEventCodec }] = await Promise.all([
    import(pathToFileURL(modulePath).href),
    import("../extension/mcp_server/event-ical.mjs"),
  ]);
  return createEventCodec(ICAL);
}

const args = {
  title: "A; title",
  description: "One\nATTENDEE: stays inside the description",
  location: "Place, Building",
  mcpTag: "test-event",
  allDay: true,
  date: "2026-11-08",
};

test("Thunderbird ICAL emits all-day dates, literal text and a stable MCP identifier", options, async () => {
  const codec = await fixture();
  const id = codec.idForTag(args.mcpTag);
  const ical = codec.create(id, args, codec.validate(args, true));
  const { event } = codec.inspect(ical, id);
  assert.equal(event.getFirstPropertyValue("summary"), "A; title (mcp:test-event)");
  assert.equal(event.getFirstPropertyValue("description"), args.description);
  assert.equal(event.getFirstPropertyValue("x-mcp-tag"), args.mcpTag);
  assert.equal(event.getFirstPropertyValue("dtstart").isDate, true);
  assert.match(ical, /DTEND;VALUE=DATE:20261109/);
  assert.equal(event.hasProperty("attendee"), false);
});

test("Thunderbird ICAL preserves unchanged properties and validates snapshots and explicit offsets", options, async () => {
  const codec = await fixture();
  const id = codec.idForTag(args.mcpTag);
  const original = codec.create(id, args, codec.validate(args, true))
    .replace("END:VEVENT", "CATEGORIES:unchanged\r\nEND:VEVENT");
  const updatedArgs = { title: "Edited", allDay: false,
    start: "2026-11-08T09:00:00-05:00", end: "2026-11-08T10:00:00-05:00" };
  const updated = codec.update(id, original, updatedArgs, codec.validate(updatedArgs, false));
  assert.match(updated, /DTSTART:20261108T140000Z/);
  assert.match(updated, /CATEGORIES:unchanged/);
  assert.match(updated, /SUMMARY:Edited \(mcp:test-event\)/);
  assert.equal(codec.sameSnapshot(original, updated, id), false);
  assert.equal(codec.sameSnapshot(original, original.replace(/DTSTAMP:\d+T\d+Z/, "DTSTAMP:20261108T120000Z"), id), true);
  assert.equal(codec.matchesWrite(updated, updated, id), true);
  assert.throws(() => codec.validate({ ...updatedArgs, start: "2026-11-08T09:00" }, false), /offset/);
  assert.throws(() => codec.validate({ ...args, date: "2026-02-30" }, true), /All-day events/);
});

test("Thunderbird ICAL refuses meeting, recurrence, and external-alarm edits", options, async () => {
  const codec = await fixture();
  const id = codec.idForTag(args.mcpTag);
  const ical = codec.create(id, args, codec.validate(args, true));
  for (const field of [
    "ATTENDEE:mailto:fixture@example.test",
    "ORGANIZER:mailto:fixture@example.test",
    "RRULE:FREQ=WEEKLY",
    "X-EAS-MEETINGSTATUS:1",
    "BEGIN:VALARM\r\nACTION:EMAIL\r\nTRIGGER:-PT10M\r\nEND:VALARM",
  ]) {
    assert.throws(() => codec.inspect(ical.replace("END:VEVENT", `${field}\r\nEND:VEVENT`), id),
      /participants or recurrence|external effects/);
  }
});
