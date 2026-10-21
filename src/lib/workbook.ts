export type RoomRecord = {
  room: string;
  students: string[];
  teachers: string[];
};

export type WorkbookSheet = {
  name: string;
  rows: unknown[][];
};

const ROOM_PATTERN = /^ม\.\s*\d+\s*\/\s*\d+$/;
const TEACHER_SPLIT_PATTERN =
  /(?:^|\s)[0-9๐-๙]+\s*[.)]\s*|(?:^|\s)[0-9๐-๙]+\s+(?=[\u0E00-\u0E7FA-Za-z])/;

type StudentNameColumn = {
  index: number;
  kind: "full" | "prefix" | "given" | "surname";
};

export function cleanCell(value: unknown): string {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

export function normalizeRoom(value: unknown): string {
  return cleanCell(value).replace(/\s+/g, "");
}

export function getRoomLevel(room: string): string {
  return /^ม\.(\d+)\//.exec(room)?.[1] ?? "";
}

function roomSortKey(room: string): number {
  const match = /^ม\.(\d+)\/(\d+)$/.exec(room);
  return match ? Number(match[1]) * 1000 + Number(match[2]) : Number.MAX_SAFE_INTEGER;
}

function splitTeachers(value: string): string[] {
  return value
    .split(TEACHER_SPLIT_PATTERN)
    .map(cleanCell)
    .filter(Boolean);
}

function studentNameColumns(sheet: WorkbookSheet): StudentNameColumn[] {
  const roomHeader = /^(?:ชั้น(?:เรียน)?|ระดับชั้น|ห้อง|class|room)$/i;

  for (const row of sheet.rows.slice(0, 20)) {
    const hasRoomHeader = row.some((cell) => roomHeader.test(cleanCell(cell)));
    if (!hasRoomHeader) continue;

    const columns = row.flatMap((cell, index): StudentNameColumn[] => {
      const header = cleanCell(cell).toLocaleLowerCase().replace(/\s+/g, "");
      if (/^ชื่อ[-–]?นาม?สกุล$/.test(header)) return [{ index, kind: "full" }];
      if (/นามสกุล|สกุล/.test(header)) return [{ index, kind: "surname" }];
      if (/คำนำหน้า|คำนำ/.test(header)) return [{ index, kind: "prefix" }];
      if (/ชื่อ/.test(header)) return [{ index, kind: "given" }];
      return [];
    });

    if (columns.length > 0) return columns;
  }

  return [];
}

export function buildRoomRecords(sheets: WorkbookSheet[]): RoomRecord[] {
  const records = new Map<string, RoomRecord>();
  const getRecord = (room: string): RoomRecord => {
    let record = records.get(room);
    if (!record) {
      record = { room, students: [], teachers: [] };
      records.set(room, record);
    }
    return record;
  };

  for (const sheet of sheets) {
    const isTeacherSheet =
      /ครู/.test(sheet.name) ||
      sheet.rows.some((row) =>
        row.some((cell) => /ครูที่ปรึกษา/.test(cleanCell(cell))),
      );
    const nameColumns = isTeacherSheet ? [] : studentNameColumns(sheet);

    for (const row of sheet.rows) {
      const roomIndex = row.findIndex((cell) => ROOM_PATTERN.test(cleanCell(cell)));
      if (roomIndex < 0) continue;
      const roomCell = row[roomIndex];
      const room = normalizeRoom(roomCell);
      const record = getRecord(room);

      if (isTeacherSheet) {
        const teachers = row
          .slice(roomIndex + 1)
          .map(cleanCell)
          .filter(Boolean)
          .flatMap(splitTeachers);
        for (const teacher of teachers) {
          if (!record.teachers.includes(teacher)) record.teachers.push(teacher);
        }
        continue;
      }

      if (nameColumns.length > 0) {
        const fullNameColumns = nameColumns.filter((column) => column.kind === "full");
        const columnsToUse =
          fullNameColumns.length > 0
            ? fullNameColumns
            : nameColumns.filter((column) => column.kind !== "full");
        const studentName = columnsToUse
          .map((column) => cleanCell(row[column.index]))
          .filter(Boolean)
          .join(" ");
        if (studentName) record.students.push(studentName);
        continue;
      }

      const cellsBeforeRoom = row
        .slice(0, roomIndex)
        .map(cleanCell)
        .filter((cell) => cell && !/^\d+$/.test(cell));
      const studentName = cellsBeforeRoom.at(-1);
      if (studentName) record.students.push(studentName);
    }
  }

  return [...records.values()].sort(
    (first, second) => roomSortKey(first.room) - roomSortKey(second.room),
  );
}
