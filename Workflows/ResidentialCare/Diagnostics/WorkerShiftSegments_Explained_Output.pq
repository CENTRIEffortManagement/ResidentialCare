// Query: WorkerShiftSegments_Explained_Output
// Purpose: Expose the exact original column layout from the separate inspection query.
let
    Output = Table.SelectColumns(
        WorkerShiftSegments_Explained,
        {
            "EmployeeID", "Facility-Abbrev", "Role", "Name",
            "Date", "Shift", "Week", "Day", "ShiftStart", "ShiftEnd",
            "AvailableHours", "FullShift", "SegmentsValid", "Disjoint"
        }
    )
in
    Table.Buffer(Output)
