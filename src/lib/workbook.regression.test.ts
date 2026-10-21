import { describe, expect, it } from "vitest";
import { buildRoomRecords } from "./workbook";

describe("workbook parser regressions", () => {
  it("skips Thai-digit student numbers without headers", () => {
    expect(buildRoomRecords([{ name: "students", rows: [["Student", "๒", "ม.2/1"]] }])).toEqual([
      { room: "ม.2/1", students: ["Student"], teachers: [] },
    ]);
  });

  it("reads advisor names when the room column is last", () => {
    expect(buildRoomRecords([{ name: "ครู", rows: [["ครูที่ปรึกษา", "ห้องเรียน"], ["ครู ก", "ม.1/2"]] }])).toEqual([
      { room: "ม.1/2", students: [], teachers: ["ครู ก"] },
    ]);
  });

  it("does not include a parent's name in the student name", () => {
    const records = buildRoomRecords([{ name: "students", rows: [
      ["เลขที่", "ชื่อ", "นามสกุล", "ชื่อบิดา", "ห้องเรียน"],
      [1, "นักเรียน ก", "ใจดี", "คุณพ่อ ก", "ม.1/2"],
    ] }]);
    expect(records[0].students).toEqual(["นักเรียน ก ใจดี"]);
  });

  it("recognizes ห้องเรียน as a room column header", () => {
    const records = buildRoomRecords([{ name: "students", rows: [
      ["เลขที่", "ชื่อ", "เพศ", "ห้องเรียน"],
      [1, "นักเรียน ก", "หญิง", "ม.1/2"],
    ] }]);
    expect(records[0].students).toEqual(["นักเรียน ก"]);
  });
});
