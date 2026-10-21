import { describe, expect, it } from "vitest";
import * as XLSX from "@e965/xlsx";
import { buildRoomRecords, getRoomLevel, normalizeRoom } from "./workbook";

describe("buildRoomRecords", () => {
  it("combines student prefix, given name, and surname from separate columns", () => {
    const records = buildRoomRecords([
      {
        name: "นร.ม.1",
        rows: [
          ["เลขที่", "คำนำหน้า", "ชื่อ", "นามสกุล", "เพศ", "ชั้น"],
          [1, "เด็กชาย", "สมชาย", "ใจดี", "ชาย", "ม.1/2"],
          [2, "เด็กหญิง", "สุดา", "รักเรียน", "หญิง", "ม.1/2"],
        ],
      },
    ]);

    expect(records).toEqual([
      {
        room: "ม.1/2",
        students: ["เด็กชาย สมชาย ใจดี", "เด็กหญิง สุดา รักเรียน"],
        teachers: [],
      },
    ]);
  });

  it("collects numbered advisors from separate columns and repeated class rows", () => {
    const records = buildRoomRecords([
      {
        name: "นร.ม.1",
        rows: [
          ["เลขที่", "ชื่อ - นามสกุล", "ชั้น"],
          [1, "นักเรียน ก", "ม.1/2"],
        ],
      },
      {
        name: "ครูม.1",
        rows: [
          ["ชั้น", "ครูที่ปรึกษาคนที่ 1", "ครูที่ปรึกษาคนที่ 2"],
          ["ม.1/2", "๑. ครู ก", "๒. ครู ข"],
          ["ม.1/2", "", "ครู ค"],
        ],
      },
    ]);

    expect(records).toEqual([
      {
        room: "ม.1/2",
        students: ["นักเรียน ก"],
        teachers: ["ครู ก", "ครู ข", "ครู ค"],
      },
    ]);
  });

  it("uses a single full-name column when available", () => {
    const records = buildRoomRecords([
      {
        name: "Students",
        rows: [
          ["เลขที่", "ชื่อ-นามสกุล", "ห้อง"],
          [1, "นางสาว ดวงใจ แสนดี", "ม.2/1"],
        ],
      },
    ]);

    expect(records[0].students).toEqual(["นางสาว ดวงใจ แสนดี"]);
  });

  it("reads actual Excel workbook data into room records", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["เลขที่", "ชื่อ", "ชั้น"],
        [1, "นักเรียน ก", "ม.1/2"],
      ]),
      "นร.ม.1",
    );
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([
        ["ชั้น", "ครูที่ปรึกษา"],
        ["ม.1/2", "ครู ก"],
      ]),
      "ครูม.1",
    );
    const parsed = XLSX.read(XLSX.write(workbook, { type: "array" }), {
      type: "array",
    });
    const sheets = parsed.SheetNames.map((name) => ({
      name,
      rows: XLSX.utils.sheet_to_json<unknown[]>(parsed.Sheets[name], {
        header: 1,
        defval: "",
        raw: false,
      }),
    }));

    expect(buildRoomRecords(sheets)).toEqual([
      {
        room: "ม.1/2",
        students: ["นักเรียน ก"],
        teachers: ["ครู ก"],
      },
    ]);
  });

  it("joins student and teacher sheets by class and sorts rooms naturally", () => {
    const records = buildRoomRecords([
      {
        name: "นร.ม.1",
        rows: [
          ["เลขที่", "ชื่อ - นามสกุล", "ชั้น"],
          [1, "เด็กชาย ก", "ม.1 / 2"],
          [2, "เด็กหญิง ข", "ม.1/2"],
          [3, "เด็กชาย ค", "ม.1/10"],
        ],
      },
      {
        name: "ครูม.1",
        rows: [
          ["ชั้น", "ครูที่ปรึกษา"],
          ["ม.1/2", "1. ครู ก 2. ครู ข"],
          ["ม.1/10", "1 ครู ค"],
        ],
      },
    ]);

    expect(records).toEqual([
      {
        room: "ม.1/2",
        students: ["เด็กชาย ก", "เด็กหญิง ข"],
        teachers: ["ครู ก", "ครู ข"],
      },
      {
        room: "ม.1/10",
        students: ["เด็กชาย ค"],
        teachers: ["ครู ค"],
      },
    ]);
  });

  it("keeps student-only rooms and skips rows without a valid room", () => {
    const records = buildRoomRecords([
      {
        name: "นักเรียน",
        rows: [
          ["เลขที่", "ชื่อ", "ชั้น"],
          [1, "นักเรียน ก", "ม.2/1"],
          [2, "ไม่มีห้อง", ""],
        ],
      },
    ]);

    expect(records).toEqual([
      { room: "ม.2/1", students: ["นักเรียน ก"], teachers: [] },
    ]);
  });
});

describe("room helpers", () => {
  it("normalizes room spacing and reads its level", () => {
    expect(normalizeRoom(" ม. 3 / 12 ")).toBe("ม.3/12");
    expect(getRoomLevel("ม.3/12")).toBe("3");
    expect(getRoomLevel("invalid")).toBe("");
  });
});
