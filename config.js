/*
 * Nexl Check — configuration.
 * Edit this file to change which tabs are checked or which columns are compared.
 *
 * Column detection: each field lists header patterns (matched against row 1, case-insensitive,
 * first matching column from the left wins). You can force a column by letter with `col: "L"`.
 */
window.NEXL_CONFIG = {
  // Nexl region filter used for the Active / Completed instruction lists: "pe", "durban", "ct" or "all".
  region: "pe",

  // Auto refresh interval while the panel is open (minutes).
  refreshMinutes: 5,

  // Name of the tab the add-in writes the discrepancy list to (created if missing).
  checkTabName: "NEXL CHECK",

  // On-sheet indicators (written at the end of each checked tab; your own cells are never changed).
  statusColumns: { stepHeader: "NEXL STEP", alertHeader: "NEXL ALERT" },
  // Raise "allocated but not started" when a driver has had the job this long with no pick-up entry.
  notStartedMinutes: 30,
  // Raise "stuck" when a truck reached a stop this long ago and hasn't reached the next one.
  stuckMinutes: 120,

  // Safety checks
  stackDatesTab: "STACK DATES",          // vessel cutoffs (NCT + PECT blocks)
  transportersTab: "DATA - TRANSPORTER", // GENSET YES/NO per transporter (trailing spaces in the tab name are fine)
  cutoffWarnHours: 12,                   // amber when a vessel cutoff is this close (red under 3h)
  pingSilentMinutes: 30,                 // "phone silent" while a truck is between stops
  whatsAppTab: "PE CITRUS",              // WhatsApp updates are built from every load on this tab

  // Header row number on every checked tab.
  headerRow: 1,

  // Max columns read per tab (A..AD).
  maxColumns: 30,

  // Fields understood by the matcher. Patterns are regexes tested against the header text.
  fields: {
    instruction: [/^INSTRUC/i],
    customer:    [/^CUSTOMER$/i],
    container:   [/^CONTAINER$/i, /^CONTAINER NO\.?$/i, /^CONTAINER NUMBER$/i],
    seal:        [/^SEAL( NO\.?)?$/i],
    booking:     [/^BOOKING (REF|NO\.?|REFERENCE)$/i],
    loadRef:     [/^LOAD REF$/i],
    vessel:      [/^VESSEL$/i],
    transporter: [/^TRANSPORTER$/i],
    driver:      [/^DRIVER( NAME)?$/i],
    genset:      [/^GENSET REQUIRED$/i],
    equipment:   [/^EQUIPMENT$/i, /^ISO$/i],
    navis:       [/^NAVIS CHECK$/i],
    comment:     [/^COMMENTS?$/i, /^REMARKS?$/i, /^CONTROLLER COMMENTS?$/i],
    tare:        [/^TARE$/i],
    loadDate:    [/LOAD DATE$/i, /^LOADING DATE$/i, /^DATE$/i],
  },

  // Tabs to check. `compare` lists the fields compared against Nexl on that tab.
  // `overrides` lets you pin a field to a column letter when the header is unreliable.
  // Short names used on the sheet that can't be worked out automatically: "SHEET NAME": ["NAME IN NEXL", ...]
  nameAliases: {
    FL4U: ["FREIGHT LOGISTICS 4U", "FREIGHT LOGISTICS FOR YOU"],
  },

  tabs: [
    {
      name: "PE CITRUS",
      cutoff: "reefer",
      slip: "comment", // port slip uploaded -> offer STACKED in COMMENT
      compare: ["seal", "booking", "loadRef", "vessel", "transporter", "driver", "customer"],
    },
    {
      name: "EXPORTS P.E",
      cutoff: "auto", // reefer cutoff when EQUIPMENT says RH/reefer, otherwise dry
      slip: "seal", // this tab writes STACKED in the SEAL column
      planned: true, // container filled in but no DRIVER yet = planned collection: no alerts until a driver is added
      // The SEAL column on this tab holds stack status / appointment times, so seal is not compared.
      compare: ["booking", "vessel", "driver", "customer"],
    },
    {
      name: "IMPORTS P.E",
      slip: "comment", // port slip on an import = container COLLECTED from the port
      planned: true, // container filled in but no DRIVER yet = planned collection: no alerts until a driver is added
      compare: ["vessel", "driver", "customer"],
    },
    {
      name: "PLUGGED IN - GROUNDED",
      cutoff: "reefer",
      compare: ["seal", "booking", "loadRef", "vessel", "driver", "customer"],
    },
  ],
};
