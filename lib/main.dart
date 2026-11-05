import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:path/path.dart' as p;
import 'package:file_saver/file_saver.dart';
import 'package:sqflite/sqflite.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await AttendanceDatabase.instance.database;
  runApp(const MyApp());
}

enum AttendanceStatus { present, absent, od, pending }

extension AttendanceStatusLabel on AttendanceStatus {
  String get label => name.toUpperCase();
}

class Student {
  final String id;
  final String className;
  final String name;
  AttendanceStatus status;

  Student({
    required this.id,
    required this.className,
    required this.name,
    this.status = AttendanceStatus.pending,
  });

  Student copyWith({String? name, AttendanceStatus? status}) => Student(
    id: id,
    className: className,
    name: name ?? this.name,
    status: status ?? this.status,
  );
}

class AttendanceRecord {
  final String studentName;
  final String date;
  final AttendanceStatus status;

  const AttendanceRecord({
    required this.studentName,
    required this.date,
    required this.status,
  });
}

class AttendanceDatabase {
  AttendanceDatabase._();
  static final AttendanceDatabase instance = AttendanceDatabase._();
  Database? _database;

  Future<Database> get database async {
    if (_database != null) return _database!;
    final databasesPath = await getDatabasesPath();
    _database = await openDatabase(
      p.join(databasesPath, 'student_attendance.db'),
      version: 1,
      onCreate: (db, version) async {
        await db.execute('''
          CREATE TABLE attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_name TEXT NOT NULL,
            student_id TEXT NOT NULL,
            student_name TEXT NOT NULL,
            attendance_date TEXT NOT NULL,
            status TEXT NOT NULL,
            UNIQUE(class_name, student_id, attendance_date)
          )
        ''');
      },
    );
    return _database!;
  }

  Future<void> saveAttendance({
    required String className,
    required List<Student> students,
    required DateTime date,
  }) async {
    final db = await database;
    final dateKey = _dateKey(date);
    final batch = db.batch();
    for (final student in students) {
      batch.insert('attendance', {
        'class_name': className,
        'student_id': student.id,
        'student_name': student.name,
        'attendance_date': dateKey,
        'status': student.status.name,
      }, conflictAlgorithm: ConflictAlgorithm.replace);
    }
    await batch.commit(noResult: true);
  }

  Future<List<AttendanceRecord>> readClassAttendance(String className) async {
    final rows = await (await database).query(
      'attendance',
      where: 'class_name = ?',
      whereArgs: [className],
      orderBy: 'attendance_date DESC, student_name ASC',
    );
    return rows
        .map(
          (row) => AttendanceRecord(
            studentName: row['student_name']! as String,
            date: row['attendance_date']! as String,
            status: AttendanceStatus.values.firstWhere(
              (value) => value.name == row['status'],
              orElse: () => AttendanceStatus.pending,
            ),
          ),
        )
        .toList();
  }

  String _dateKey(DateTime date) {
    final local = date.toLocal();
    return '${local.year.toString().padLeft(4, '0')}-'
        '${local.month.toString().padLeft(2, '0')}-'
        '${local.day.toString().padLeft(2, '0')}';
  }
}

class ExcelExportService {
  static Future<String?> exportClassAttendance(
    String className,
    List<Student> fallbackStudents,
  ) async {
    final records = await AttendanceDatabase.instance.readClassAttendance(
      className,
    );
    final tableRecords = records.isEmpty
        ? fallbackStudents
              .map(
                (student) => AttendanceRecord(
                  studentName: student.name,
                  date: DateTime.now().toIso8601String().split('T').first,
                  status: student.status,
                ),
              )
              .toList()
        : records;

    final dates = tableRecords.map((record) => record.date).toSet().toList();
    final studentNames =
        tableRecords.map((record) => record.studentName).toSet().toList()
          ..sort();
    final byStudentAndDate = {
      for (final record in tableRecords)
        '${record.studentName}|${record.date}': record.status,
    };

    final workbook = Excel.createExcel();
    // Write into Excel's default sheet so the table is visible immediately
    // when the user opens the exported workbook.
    final sheet = workbook['Sheet1'];
    sheet.appendRow([
      TextCellValue('Student'),
      ...dates.map((date) => TextCellValue(date)),
    ]);
    for (final student in studentNames) {
      sheet.appendRow([
        TextCellValue(student),
        ...dates.map(
          (date) =>
              TextCellValue(byStudentAndDate['$student|$date']?.label ?? '-'),
        ),
      ]);
    }

    final bytes = workbook.save();
    if (bytes == null) throw Exception('Could not create workbook');
    final safeClass = className.replaceAll(RegExp(r'[^a-zA-Z0-9_-]+'), '_');
    return FileSaver.instance.saveAs(
      name: '${safeClass}_attendance',
      bytes: Uint8List.fromList(bytes),
      ext: 'xlsx',
      mimeType: MimeType.microsoftExcel,
    );
  }
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      theme: ThemeData(useMaterial3: true, colorSchemeSeed: Colors.blue),
      home: const HomeScreen(),
    );
  }
}

class HomeScreen extends StatefulWidget {
  const HomeScreen({super.key});

  @override
  State<HomeScreen> createState() => _HomeScreenState();
}

class _HomeScreenState extends State<HomeScreen> {
  final Map<String, List<Student>> _classStudents = {
    'Class 10-A': _makeStudents('Class 10-A', [
      'Alice Smith',
      'Bob Johnson',
      'Charlie Brown',
      'Diana Prince',
      'Eve Adams',
    ]),
    'Class 10-B': _makeStudents('Class 10-B', [
      'Frank Miller',
      'Grace Lee',
      'Henry Wilson',
      'Ivy Thomas',
    ]),
    'Class 11-A': _makeStudents('Class 11-A', [
      'Jack Brown',
      'Kate Davis',
      'Liam Moore',
      'Mia Taylor',
    ]),
  };

  String? _selectedClass = 'Class 10-A';

  static List<Student> _makeStudents(String className, List<String> names) => [
    for (var i = 0; i < names.length; i++)
      Student(id: '${className}_$i', className: className, name: names[i]),
  ];

  List<Student> get _selectedStudents => _classStudents[_selectedClass] ?? [];

  Future<void> _openAttendance() async {
    final result = await Navigator.push<List<Student>>(
      context,
      MaterialPageRoute(
        builder: (_) => SwipeCard(
          className: _selectedClass!,
          students: _selectedStudents
              .map((student) => student.copyWith())
              .toList(),
        ),
      ),
    );
    if (result != null && mounted)
      setState(() => _classStudents[_selectedClass!] = result);
  }

  Future<void> _viewAttendance() async {
    await Navigator.push(
      context,
      MaterialPageRoute(
        builder: (_) => AttendanceHistoryScreen(className: _selectedClass!),
      ),
    );
  }

  Future<void> _exportExcel() async {
    try {
      final path = await ExcelExportService.exportClassAttendance(
        _selectedClass!,
        _selectedStudents,
      );
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text(
            path == null
                ? 'Export cancelled'
                : 'Excel saved in the selected device folder',
          ),
        ),
      );
    } catch (error) {
      if (mounted)
        ScaffoldMessenger.of(
          context,
        ).showSnackBar(SnackBar(content: Text('Export failed: $error')));
    }
  }

  Future<void> _editStudents() async {
    final result = await Navigator.push<List<Student>>(
      context,
      MaterialPageRoute(
        builder: (_) => StudentEditorScreen(
          className: _selectedClass!,
          students: _selectedStudents
              .map((student) => student.copyWith())
              .toList(),
        ),
      ),
    );
    if (result != null && mounted)
      setState(() => _classStudents[_selectedClass!] = result);
  }

  Future<void> _addClass() async {
    final result = await showDialog<(String, List<String>)>(
      context: context,
      builder: (_) => const AddClassDialog(),
    );
    if (result == null || result.$1.trim().isEmpty || result.$2.isEmpty) return;
    final className = result.$1.trim();
    if (_classStudents.containsKey(className)) return;
    setState(() {
      _classStudents[className] = _makeStudents(className, result.$2);
      _selectedClass = className;
    });
  }

  @override
  Widget build(BuildContext context) {
    final classes = _classStudents.keys.toList();
    return Scaffold(
      backgroundColor: Colors.grey.shade200,
      appBar: AppBar(title: const Text('Student Attendance')),
      body: Padding(
        padding: const EdgeInsets.all(16),
        child: Column(
          children: [
            Card(
              child: Padding(
                padding: const EdgeInsets.fromLTRB(16, 16, 16, 20),
                child: Column(
                  children: [
                    const Text(
                      'Attendance',
                      style: TextStyle(
                        fontSize: 22,
                        fontWeight: FontWeight.bold,
                      ),
                    ),
                    const SizedBox(height: 12),
                    Center(
                      child: DropdownButton<String>(
                        value: _selectedClass,
                        alignment: AlignmentDirectional.center,
                        items: classes
                            .map(
                              (name) => DropdownMenuItem(
                                value: name,
                                child: Text(name),
                              ),
                            )
                            .toList(),
                        onChanged: (value) =>
                            setState(() => _selectedClass = value),
                      ),
                    ),
                    const SizedBox(height: 18),
                    SizedBox(
                      width: double.infinity,
                      child: ElevatedButton(
                        onPressed: _openAttendance,
                        child: const Text('Take Attendance'),
                      ),
                    ),
                    const SizedBox(height: 10),
                    SizedBox(
                      width: double.infinity,
                      child: OutlinedButton(
                        onPressed: _viewAttendance,
                        child: const Text('View Attendance'),
                      ),
                    ),
                    const SizedBox(height: 10),
                    SizedBox(
                      width: double.infinity,
                      child: OutlinedButton.icon(
                        onPressed: _exportExcel,
                        icon: const Icon(Icons.table_view),
                        label: const Text('Export as Excel'),
                      ),
                    ),
                    const SizedBox(height: 10),
                    SizedBox(
                      width: double.infinity,
                      child: OutlinedButton.icon(
                        onPressed: _editStudents,
                        icon: const Icon(Icons.people_alt_outlined),
                        label: const Text('Edit Student'),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton.extended(
        onPressed: _addClass,
        icon: const Icon(Icons.add),
        label: const Text('Add class'),
      ),
    );
  }
}

class AddClassDialog extends StatefulWidget {
  const AddClassDialog({super.key});
  @override
  State<AddClassDialog> createState() => _AddClassDialogState();
}

class _AddClassDialogState extends State<AddClassDialog> {
  final classController = TextEditingController();
  final studentsController = TextEditingController();

  @override
  void dispose() {
    classController.dispose();
    studentsController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
    title: const Text('Add new class'),
    content: Column(
      mainAxisSize: MainAxisSize.min,
      children: [
        TextField(
          controller: classController,
          autofocus: true,
          decoration: const InputDecoration(labelText: 'Class name'),
        ),
        TextField(
          controller: studentsController,
          maxLines: 3,
          decoration: const InputDecoration(
            labelText: 'Student names, one per line',
          ),
        ),
      ],
    ),
    actions: [
      TextButton(
        onPressed: () => Navigator.pop(context),
        child: const Text('Cancel'),
      ),
      FilledButton(
        onPressed: () {
          final names = studentsController.text
              .split('\n')
              .map((name) => name.trim())
              .where((name) => name.isNotEmpty)
              .toList();
          Navigator.pop(context, (classController.text, names));
        },
        child: const Text('Add'),
      ),
    ],
  );
}

class SwipeCard extends StatefulWidget {
  final String className;
  final List<Student> students;
  const SwipeCard({super.key, required this.className, required this.students});
  @override
  State<SwipeCard> createState() => _SwipeCardState();
}

class _SwipeCardState extends State<SwipeCard> {
  late final List<Student> students = widget.students;
  int currentIndex = 0;
  double x = 0, y = 0;
  String swipeDirection = '';
  Color glowColor = Colors.transparent;
  bool isEditingPrevious = false;
  bool isSaving = false;

  void _processSwipe(AttendanceStatus status) {
    if (currentIndex >= students.length) return;
    setState(() {
      students[currentIndex].status = status;
      currentIndex++;
      isEditingPrevious = false;
      _resetCardPosition();
    });
    if (currentIndex >= students.length) _saveAndShowSummary();
  }

  void _resetCardPosition() {
    x = 0;
    y = 0;
    swipeDirection = '';
    glowColor = Colors.transparent;
  }

  Future<void> _saveAndShowSummary() async {
    setState(() => isSaving = true);
    await AttendanceDatabase.instance.saveAttendance(
      className: widget.className,
      students: students,
      date: DateTime.now(),
    );
    if (!mounted) return;
    setState(() => isSaving = false);
    final result = await Navigator.push<List<Student>>(
      context,
      MaterialPageRoute(
        builder: (_) => AttendanceListScreen(
          className: widget.className,
          students: students,
        ),
      ),
    );
    if (mounted) Navigator.pop(context, result ?? students);
  }

  void _updateDirection() {
    if (y < -80) {
      swipeDirection = 'PRESENT';
      glowColor = Colors.green;
    } else if (y > 80) {
      swipeDirection = 'ABSENT';
      glowColor = Colors.red;
    } else if (x < -80) {
      swipeDirection = 'OD';
      glowColor = Colors.amber;
    } else if (x > 80 && currentIndex > 0) {
      swipeDirection = 'EDIT PREVIOUS';
      glowColor = Colors.blue;
    } else {
      swipeDirection = '';
      glowColor = Colors.transparent;
    }
  }

  void _finishPan() {
    if (swipeDirection == 'PRESENT')
      _processSwipe(AttendanceStatus.present);
    else if (swipeDirection == 'ABSENT')
      _processSwipe(AttendanceStatus.absent);
    else if (swipeDirection == 'OD')
      _processSwipe(AttendanceStatus.od);
    else if (swipeDirection == 'EDIT PREVIOUS' && currentIndex > 0) {
      setState(() {
        currentIndex--;
        isEditingPrevious = true;
        _resetCardPosition();
      });
    } else
      setState(_resetCardPosition);
  }

  @override
  Widget build(BuildContext context) {
    final hasCard = currentIndex < students.length;
    return Scaffold(
      appBar: AppBar(title: Text('${widget.className} Attendance')),
      body: Center(
        child: isSaving
            ? const CircularProgressIndicator()
            : hasCard
            ? GestureDetector(
                onPanUpdate: (details) => setState(() {
                  x += details.delta.dx;
                  y += details.delta.dy;
                  _updateDirection();
                }),
                onPanEnd: (_) => _finishPan(),
                child: Transform.translate(
                  offset: Offset(x, y),
                  child: Transform.rotate(
                    angle: x / 500,
                    child: _buildCard(students[currentIndex]),
                  ),
                ),
              )
            : const Text('All students processed!'),
      ),
    );
  }

  Widget _buildCard(Student student) => Stack(
    alignment: Alignment.center,
    children: [
      AnimatedContainer(
        duration: const Duration(milliseconds: 200),
        width: 300,
        height: 400,
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(20),
          boxShadow: [
            BoxShadow(
              color: glowColor.withOpacity(.8),
              blurRadius: 40,
              spreadRadius: 8,
            ),
          ],
        ),
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Text(
              student.name,
              style: const TextStyle(fontSize: 35, fontWeight: FontWeight.bold),
              textAlign: TextAlign.center,
            ),
            const SizedBox(height: 20),
            Text(
              'Status: ${student.status.label}',
              style: TextStyle(fontSize: 20, color: Colors.grey[600]),
            ),
          ],
        ),
      ),
      Positioned(
        top: 20,
        child: AnimatedOpacity(
          opacity: swipeDirection.isEmpty ? 0 : 1,
          duration: const Duration(milliseconds: 150),
          child: Container(
            padding: const EdgeInsets.symmetric(horizontal: 15, vertical: 8),
            decoration: BoxDecoration(
              color: glowColor,
              borderRadius: BorderRadius.circular(10),
            ),
            child: Text(
              swipeDirection,
              style: const TextStyle(
                color: Colors.white,
                fontSize: 20,
                fontWeight: FontWeight.bold,
              ),
            ),
          ),
        ),
      ),
      if (isEditingPrevious) ...[
        Positioned(
          bottom: 20,
          left: 20,
          child: _actionButton(
            Colors.green,
            Icons.check,
            () => _processSwipe(AttendanceStatus.present),
          ),
        ),
        Positioned(
          bottom: 20,
          right: 20,
          child: _actionButton(
            Colors.red,
            Icons.close,
            () => _processSwipe(AttendanceStatus.absent),
          ),
        ),
      ],
    ],
  );

  Widget _actionButton(Color color, IconData icon, VoidCallback onTap) =>
      GestureDetector(
        onTap: onTap,
        child: CircleAvatar(
          radius: 30,
          backgroundColor: color,
          child: Icon(icon, color: Colors.white, size: 40),
        ),
      );
}

class AttendanceListScreen extends StatefulWidget {
  final String className;
  final List<Student> students;
  const AttendanceListScreen({
    super.key,
    required this.className,
    required this.students,
  });
  @override
  State<AttendanceListScreen> createState() => _AttendanceListScreenState();
}

class _AttendanceListScreenState extends State<AttendanceListScreen> {
  late final List<Student> students = widget.students;

  AttendanceStatus _nextStatus(AttendanceStatus current) {
    switch (current) {
      case AttendanceStatus.present:
        return AttendanceStatus.absent;
      case AttendanceStatus.absent:
        return AttendanceStatus.od;
      case AttendanceStatus.od:
        return AttendanceStatus.present;
      case AttendanceStatus.pending:
        return AttendanceStatus.present;
    }
  }

  Future<void> _save() => AttendanceDatabase.instance.saveAttendance(
    className: widget.className,
    students: students,
    date: DateTime.now(),
  );

  @override
  Widget build(BuildContext context) => Scaffold(
    appBar: AppBar(title: const Text('Attendance Summary')),
    body: Column(
      children: [
        const Padding(
          padding: EdgeInsets.all(12),
          child: Text(
            'Double-tap any status to cycle Present → Absent → OD → Present.',
            textAlign: TextAlign.center,
          ),
        ),
        Expanded(
          child: ListView.builder(
            itemCount: students.length,
            itemBuilder: (context, index) {
              final student = students[index];
              return Card(
                margin: const EdgeInsets.symmetric(horizontal: 16, vertical: 6),
                child: ListTile(
                  title: Text(student.name),
                  subtitle: Text('Double-tap status to edit'),
                  trailing: GestureDetector(
                    onDoubleTap: () async {
                      setState(
                        () => student.status = _nextStatus(student.status),
                      );
                      await _save();
                    },
                    child: Padding(
                      padding: const EdgeInsets.all(8),
                      child: Text(
                        student.status.label,
                        style: TextStyle(
                          color: _statusColor(student.status),
                          fontWeight: FontWeight.bold,
                        ),
                      ),
                    ),
                  ),
                ),
              );
            },
          ),
        ),
        Padding(
          padding: const EdgeInsets.all(16),
          child: SizedBox(
            width: double.infinity,
            child: FilledButton.icon(
              onPressed: () => Navigator.pop(context, students),
              icon: const Icon(Icons.done),
              label: const Text('Done'),
            ),
          ),
        ),
      ],
    ),
  );
}

class EditStudentDialog extends StatefulWidget {
  final String initialName;

  const EditStudentDialog({super.key, required this.initialName});

  @override
  State<EditStudentDialog> createState() => _EditStudentDialogState();
}

class _EditStudentDialogState extends State<EditStudentDialog> {
  late final TextEditingController controller;

  @override
  void initState() {
    super.initState();
    controller = TextEditingController(text: widget.initialName);
  }

  @override
  void dispose() {
    controller.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
    title: const Text('Edit student name'),
    content: TextField(
      controller: controller,
      autofocus: true,
      textInputAction: TextInputAction.done,
      decoration: const InputDecoration(labelText: 'Name'),
      onSubmitted: (value) => Navigator.pop(context, value.trim()),
    ),
    actions: [
      TextButton(
        onPressed: () => Navigator.pop(context),
        child: const Text('Cancel'),
      ),
      FilledButton(
        onPressed: () => Navigator.pop(context, controller.text.trim()),
        child: const Text('Save'),
      ),
    ],
  );
}

class StudentEditorScreen extends StatefulWidget {
  final String className;
  final List<Student> students;
  const StudentEditorScreen({
    super.key,
    required this.className,
    required this.students,
  });
  @override
  State<StudentEditorScreen> createState() => _StudentEditorScreenState();
}

class _StudentEditorScreenState extends State<StudentEditorScreen> {
  late List<Student> students = widget.students;
  int? selectedIndex;

  Future<void> _editSelected() async {
    final index = selectedIndex;
    if (index == null) return;
    final value = await showDialog<String>(
      context: context,
      builder: (_) => EditStudentDialog(initialName: students[index].name),
    );
    if (!mounted) return;
    if (value != null && value.isNotEmpty) {
      setState(() => students[index] = students[index].copyWith(name: value));
    }
  }

  void _deleteSelected() {
    final index = selectedIndex;
    if (index == null) return;
    setState(() {
      students.removeAt(index);
      selectedIndex = null;
    });
  }

  @override
  Widget build(BuildContext context) => Scaffold(
    appBar: AppBar(
      title: Text('${widget.className} Students'),
      actions: [
        IconButton(
          tooltip: 'Edit selected student',
          onPressed: _editSelected,
          icon: const Icon(Icons.edit),
        ),
        IconButton(
          tooltip: 'Delete selected student',
          onPressed: _deleteSelected,
          icon: const Icon(Icons.delete),
        ),
      ],
    ),
    body: students.isEmpty
        ? const Center(child: Text('No students in this class.'))
        : ListView.builder(
            itemCount: students.length,
            itemBuilder: (context, index) => ListTile(
              selected: selectedIndex == index,
              onTap: () => setState(() => selectedIndex = index),
              leading: Icon(
                selectedIndex == index
                    ? Icons.radio_button_checked
                    : Icons.radio_button_unchecked,
              ),
              title: Text(students[index].name),
              subtitle: Text('ID: ${students[index].id}'),
            ),
          ),
    floatingActionButton: FloatingActionButton(
      onPressed: () => Navigator.pop(context, students),
      child: const Icon(Icons.done),
    ),
  );
}

class AttendanceHistoryScreen extends StatefulWidget {
  final String className;
  const AttendanceHistoryScreen({super.key, required this.className});
  @override
  State<AttendanceHistoryScreen> createState() =>
      _AttendanceHistoryScreenState();
}

class _AttendanceHistoryScreenState extends State<AttendanceHistoryScreen> {
  late Future<List<AttendanceRecord>> records;

  @override
  void initState() {
    super.initState();
    records = AttendanceDatabase.instance.readClassAttendance(widget.className);
  }

  @override
  Widget build(BuildContext context) => Scaffold(
    appBar: AppBar(title: Text('${widget.className} History')),
    body: FutureBuilder<List<AttendanceRecord>>(
      future: records,
      builder: (context, snapshot) {
        if (snapshot.connectionState != ConnectionState.done) {
          return const Center(child: CircularProgressIndicator());
        }
        if (snapshot.hasError) {
          return Center(
            child: Text('Could not load attendance: ${snapshot.error}'),
          );
        }
        final records = snapshot.data ?? [];
        if (records.isEmpty) {
          return const Center(child: Text('No attendance has been saved yet.'));
        }

        final dates = records.map((record) => record.date).toSet().toList();
        final students =
            records.map((record) => record.studentName).toSet().toList()
              ..sort();
        final byStudentAndDate = {
          for (final record in records)
            '${record.studentName}|${record.date}': record.status,
        };

        return SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          child: SingleChildScrollView(
            child: DataTable(
              columns: [
                const DataColumn(label: Text('Student')),
                ...dates.map((date) => DataColumn(label: Text(date))),
              ],
              rows: students.map((student) {
                return DataRow(
                  cells: [
                    DataCell(Text(student)),
                    ...dates.map((date) {
                      final status = byStudentAndDate['$student|$date'];
                      return DataCell(
                        Text(
                          status?.label ?? '-',
                          style: TextStyle(
                            color: status == null
                                ? Colors.grey
                                : _statusColor(status),
                            fontWeight: FontWeight.bold,
                          ),
                        ),
                      );
                    }),
                  ],
                );
              }).toList(),
            ),
          ),
        );
      },
    ),
  );
}

Color _statusColor(AttendanceStatus status) {
  switch (status) {
    case AttendanceStatus.present:
      return Colors.green;
    case AttendanceStatus.absent:
      return Colors.red;
    case AttendanceStatus.od:
      return Colors.amber.shade800;
    case AttendanceStatus.pending:
      return Colors.grey;
  }
}

// pubspec.yaml dependencies used by this file:
//   path: ^1.9.0
//   sqflite: ^2.3.3+1
//   file_saver: ^0.2.14
//   excel: ^4.0.6
// file_saver opens the platform's user-visible save flow, so the user can
// choose Downloads or another normal folder in the device file manager.
