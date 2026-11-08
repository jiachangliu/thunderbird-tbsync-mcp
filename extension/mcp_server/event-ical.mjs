const TAG = /^[a-z0-9][a-z0-9-]{2,63}$/;
const TAG_SUFFIX = / \(mcp:([a-z0-9][a-z0-9-]{2,63})\)$/;
const DATE = /^\d{4}-\d{2}-\d{2}$/;

function validateDate(value) {
  if (typeof value !== "string" || !DATE.test(value) || value.startsWith("0000")) return false;
  const date = new Date(`${value}T00:00:00Z`);
  return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

export function parseTimestamp(value) {
  const pattern = /^\d{4}-\d{2}-\d{2}T(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?(?:Z|[+-](?:[01]\d|2[0-3]):[0-5]\d)$/;
  if (typeof value !== "string" || !pattern.test(value) || !validateDate(value.slice(0, 10))) return NaN;
  const instant = new Date(value);
  const year = instant.getUTCFullYear();
  return year >= 1 && year <= 9999 ? instant.getTime() : NaN;
}

function textField(args, field, max, required = false) {
  const value = args[field];
  if (value === undefined && !required) return;
  if (typeof value !== "string" || value.length > max || (required && !value.trim()) ||
      /[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f]/.test(value)) {
    throw new Error(`${field} must be valid text (up to ${max} characters).`);
  }
}

function assertTag(tag) {
  if (typeof tag !== "string" || !TAG.test(tag)) {
    throw new Error("mcpTag must contain 3–64 lowercase letters, digits or hyphens and start with a letter or digit.");
  }
}

export function createEventCodec(ICAL) {
  function inspect(ical, id) {
    const root = new ICAL.Component(ICAL.parse(ical));
    const components = root.getAllSubcomponents();
    const events = root.getAllSubcomponents("vevent");
    if (root.name !== "vcalendar" || root.hasProperty("method") || events.length !== 1 ||
        components.some(component => !["vevent", "vtimezone"].includes(component.name))) {
      throw new Error("Only a single standalone event can be modified.");
    }
    const event = events[0];
    if (event.getFirstPropertyValue("uid") !== id) throw new Error("Event UID does not match the requested id.");
    if (["attendee", "organizer", "rrule", "rdate", "exrule", "exdate", "recurrence-id"]
      .some(name => event.hasProperty(name)) ||
        Number(event.getFirstPropertyValue("x-eas-meetingstatus") ?? 0) !== 0) {
      throw new Error("Events with scheduling participants or recurrence cannot be changed through this tool.");
    }
    if (event.getAllSubcomponents().some(component => component.name !== "valarm" ||
      !["DISPLAY", "AUDIO"].includes(String(component.getFirstPropertyValue("action")).toUpperCase()))) {
      throw new Error("This event contains an alarm or component that could have external effects.");
    }
    return { root, event };
  }

  function timing(args, required) {
    if (!required && !["allDay", "date", "endDate", "start", "end"].some(field => args[field] !== undefined)) return;
    if (args.allDay === true) {
      const start = args.date;
      if (!validateDate(start) || args.start !== undefined || args.end !== undefined) {
        throw new Error("All-day events need date (YYYY-MM-DD) and optional exclusive endDate.");
      }
      const end = args.endDate ?? new Date(Date.parse(`${start}T00:00:00Z`) + 86400000).toISOString().slice(0, 10);
      if (!validateDate(end) || end <= start) throw new Error("endDate must be a later exclusive date.");
      return { start: ICAL.Time.fromDateString(start), end: ICAL.Time.fromDateString(end) };
    }
    const start = parseTimestamp(args.start);
    const end = parseTimestamp(args.end);
    if (args.allDay !== false || args.date !== undefined || args.endDate !== undefined ||
        !Number.isFinite(start) || !Number.isFinite(end) || end <= start) {
      throw new Error("Timed events need allDay=false and increasing start and end with an explicit UTC offset or Z.");
    }
    const toTime = time => ICAL.Time.fromJSDate(new Date(time), true);
    return { start: toTime(start), end: toTime(end) };
  }

  function validate(args, creating) {
    textField(args, "title", 512, creating || args.title !== undefined);
    textField(args, "description", 16384);
    textField(args, "location", 1024);
    if (creating || args.mcpTag !== undefined) assertTag(args.mcpTag);
    return timing(args, creating);
  }

  function apply(event, args, times) {
    const summary = String(event.getFirstPropertyValue("summary") ?? "");
    const priorTag = event.getFirstPropertyValue("x-mcp-tag") ?? summary.match(TAG_SUFFIX)?.[1];
    const tag = args.mcpTag ?? priorTag;
    assertTag(tag);
    if (priorTag && priorTag !== tag) throw new Error("Keep the original MCP tag when updating an event.");
    const title = (args.title ?? summary).replace(TAG_SUFFIX, "").trim();
    if (!title) throw new Error("An event title is required.");
    event.updatePropertyWithValue("summary", `${title} (mcp:${tag})`);
    event.updatePropertyWithValue("x-mcp-tag", tag);
    for (const name of ["description", "location"]) {
      if (args[name] !== undefined) event.updatePropertyWithValue(name, args[name]);
    }
    if (times) {
      event.removeAllProperties("duration");
      for (const [name, time] of [["dtstart", times.start], ["dtend", times.end]]) {
        // Replace the property so a change from zoned time to UTC or DATE
        // cannot retain the old TZID or VALUE parameter.
        event.removeAllProperties(name);
        event.addPropertyWithValue(name, time);
      }
    }
  }

  return {
    validate,
    inspect,
    idForTag(tag) {
      return `tbsync-mcp-${tag}@luthriel.dev`;
    },
    create(id, args, times) {
      const root = new ICAL.Component("vcalendar");
      root.addPropertyWithValue("version", "2.0");
      root.addPropertyWithValue("prodid", "-//Thunderbird TbSync MCP//EN");
      const event = new ICAL.Component("vevent");
      root.addSubcomponent(event);
      event.addPropertyWithValue("uid", id);
      event.addPropertyWithValue("dtstamp", ICAL.Time.fromJSDate(new Date(), true));
      apply(event, args, times);
      return root.toString();
    },
    sameSnapshot(current, expected, id) {
      const normalize = ical => {
        const { root, event } = inspect(ical, id);
        for (const name of ["dtstamp", "last-modified"]) event.removeAllProperties(name);
        return root.toString();
      };
      return normalize(current) === normalize(expected);
    },
    matchesWrite(current, expected, id) {
      const actual = inspect(current, id).event;
      const intended = inspect(expected, id).event;
      for (const name of ["summary", "description", "location", "dtstart", "dtend"]) {
        const left = actual.getFirstPropertyValue(name);
        const right = intended.getFirstPropertyValue(name);
        if (left?.isDate === false && right?.isDate === false) {
          if (left.toUnixTime() !== right.toUnixTime()) return false;
        } else if (String(left ?? "") !== String(right ?? "")) {
          return false;
        }
      }
      return true;
    },
    update(id, ical, args, times) {
      const { root, event } = inspect(ical, id);
      apply(event, args, times);
      return root.toString();
    },
  };
}
