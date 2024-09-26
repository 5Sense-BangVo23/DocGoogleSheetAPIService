# ProjectGoogleSheetV2Listener
Lớp `ProjectGoogleSheetV2Listener` xử lý các sự kiện liên quan đến việc cập nhật hoặc tạo mới các bảng tính Google Sheets cho các dự án. Lớp này tương tác với dịch vụ bảng tính để quản lý việc tạo, cập nhật và làm mới các bảng tính dựa trên các sự kiện nhận được.

## Class Structure
### Class declare
- Nhận các dịch vụ cần thiết: `SpreadsheetsService` và `ProjectSheetServiceV2`.

### Main method
- **handle(event: ProjectGoogleSheetEvent)**: Xử lý sự kiện và quyết định xem có cần tạo mới bảng tính hay cập nhật bảng tính đã tồn tại hay không.

### Secondary methods
- **createSheet(title, projectId, isUpdated, currentMonth, monthLabel, currentYear)**: Tạo một bảng tính mới và điền dữ liệu khởi đầu.
- **handleExistingSheets(sheetTitles, sheetTitle, projectId, isCreated, isUpdated, currentMonth, monthLabel, currentYear)**: Xử lý các bảng tính đã tồn tại.
- **clearSheetData(sheetTitle)**: Xóa các xử lý về màu sắc của ô dữ liệu trong bảng tính.
- **initializeOrRefreshSheetTemplate(sheetTitle, isCreated, isUpdated, projectId, currentMonth, monthLabel, currentYear)**: Khởi tạo hoặc làm mới mẫu bảng tính.
- **populateSheetWithInitialData(sheetTitle, projectId, isUpdated, currentMonth, monthLabel, currentYear)**: Điền dữ liệu khởi đầu cho bảng tính.
- **updateInitialData(sheetTitle, projectId, isUpdated, monthLabel, currentMonth, currentYear)**: Cập nhật dữ liệu cho bảng tính hiện có.
- **addInitialData(sheetTitle)**: Thêm các ô dữ liệu khởi đầu.
- **setupAdditionalData(sheetTitle)**: Thêm các ô bổ sung cho bảng tính.
- **addTitleRowAndFreezeColumns(sheetTitle)**: Thêm hàng tiêu đề và cố định cột.
- **getMonthLabel(month)**: Trả về nhãn tháng tương ứng với số tháng bằng tiếng Nhật.

## Pseudo code

```pseudo
class ProjectGoogleSheetV2Listener
    method __construct(SpreadsheetsService googleSheetService, ProjectSheetServiceV2 projectSheetService)
        this.googleSheetService = googleSheetService
        this.projectSheetService = projectSheetService

    method handle(ProjectGoogleSheetEvent event)
        sheetTitles = this.googleSheetService.getSheetTitles()
        projectId = event.project_id
        date = current date
        currentMonthInteger = date.getMonth()
        currentYearInteger = date.getYear()
        monthLabel = getMonthLabel(currentMonthInteger)

        currentMonthSheetTitle = this.googleSheetService.generateSheetTitle()

        if currentMonthSheetTitle not in sheetTitles
            createSheet(currentMonthSheetTitle, projectId, event.is_updated, currentMonthInteger, monthLabel, currentYearInteger)
        else
            handleExistingSheets(sheetTitles, currentMonthSheetTitle, projectId, event.is_created, event.is_updated, currentMonthInteger, monthLabel, currentYearInteger)

    private method createSheet(title, projectId, isUpdated, currentMonth, monthLabel, currentYear)
        this.googleSheetService.addNewSheet(title)
        populateSheetWithInitialData(title, projectId, isUpdated, currentMonth, monthLabel, currentYear)

    private method handleExistingSheets(sheetTitles, sheetTitle, projectId, isCreated, isUpdated, currentMonth, monthLabel, currentYear)
        clearSheetData(sheetTitle)

        if sheetTitle not in sheetTitles
            createSheet(sheetTitle, projectId, isUpdated, currentMonth, monthLabel, currentYear)
        else
            populateSheetWithInitialData(sheetTitle, projectId, isUpdated, currentMonth, monthLabel, currentYear)

        if isCreated or isUpdated
            initializeOrRefreshSheetTemplate(sheetTitle, isCreated, isUpdated, projectId, currentMonth, monthLabel, currentYear)

    private method clearSheetData(sheetTitle)
        this.googleSheetService.clearSheetDataProject(sheetTitle)
        this.googleSheetService.clearSheetDataCalendar(sheetTitle)

    private method initializeOrRefreshSheetTemplate(sheetTitle, isCreated, isUpdated, projectId, currentMonth, monthLabel, currentYear)
        try
            isSheetExisting = this.googleSheetService.isSheetCached(sheetTitle)

            if not isSheetExisting
                isSheetExisting = this.googleSheetService.executeWithRetry(() => this.googleSheetService.isSheetExist(sheetTitle))

            if isCreated or isUpdated
                if not isSheetExisting
                    this.googleSheetService.addNewSheet(sheetTitle)

                if isCreated
                    populateSheetWithInitialData(sheetTitle, projectId, isUpdated, currentMonth, monthLabel, currentYear)
                else if isUpdated
                    updateInitialData(sheetTitle, projectId, isUpdated, monthLabel, currentMonth, currentYear)
        catch Exception e
            throw new Exception('Error initializing or refreshing sheet template: ' + e.getMessage())

    private method populateSheetWithInitialData(sheetTitle, projectId, isUpdated, currentMonth, monthLabel, currentYear)
        if sheetTitle == this.googleSheetService.generateSheetTitle()
            addInitialData(sheetTitle)
            setupAdditionalData(sheetTitle)
            addTitleRowAndFreezeColumns(sheetTitle)
            this.projectSheetService.setupCellsAndMerge(sheetTitle, monthLabel)
            this.googleSheetService.displayMonthInSheet(sheetTitle, [[startRow: 7, startColumn: 10]], currentMonth, currentYear, "#FFF")
            this.projectSheetService.loadDataToCellFromDatabase(sheetTitle, projectId, isUpdated, currentMonth, currentYear)
        else
            throw new Exception('Sheet title is not the current month.')

    private method updateInitialData(sheetTitle, projectId, isUpdated, monthLabel, currentMonth, currentYear)
        addInitialData(sheetTitle)
        setupAdditionalData(sheetTitle)
        this.projectSheetService.loadDataToCellFromDatabase(sheetTitle, projectId, isUpdated, currentMonth, currentYear)
        addTitleRowAndFreezeColumns(sheetTitle)
        this.projectSheetService.setupCellsAndMerge(sheetTitle, monthLabel)

    private method addInitialData(sheetTitle)
        initialData = [[1, 1, '', 12, false], ...] // Thêm các dữ liệu khởi đầu
        for data in initialData
            [row, col, text, fontSize, bold] = data
            this.googleSheetService.inputTextToCell(sheetTitle, row, col, text, fontSize, bold)

    private method setupAdditionalData(sheetTitle)
        this.googleSheetService.mergeCellsInRange(sheetTitle, 4, 4, 0, 2, "", false, 12, null)
        additionalData = [[2, 9, '', 12, true], ...] // Dữ liệu bổ sung
        for data in additionalData
            [row, col, text, fontSize, bold] = data
            this.googleSheetService.inputTextToCell(sheetTitle, row, col, text, fontSize, bold)

    private method addTitleRowAndFreezeColumns(sheetTitle)
        titleProject = ['加工書№', '得意先', ...] // Tiêu đề dự án
        this.googleSheetService.addTitleProjectRow(sheetTitle, titleProject, 6, 0)
        sheetId = this.googleSheetService.getSheetIdByTitle(sheetTitle)
        this.googleSheetService.freezeColumns(sheetId, 10)

    private method getMonthLabel(month)
        monthNamesInJapanese = {1: '1月', 2: '2月', ...} // Định nghĩa tên tháng
        return monthNamesInJapanese[month]


####--------------------------------------------------------------------------------------------

#ProjectSheetServiceV2

`ProjectSheetServiceV2` là một lớp trong namespace `App\Services` có trách nhiệm xử lý và cập nhật dữ liệu của các dự án vào Google Sheets. Lớp này tương tác với cơ sở dữ liệu và Google Sheets API để tải dữ liệu dự án, ghi lại các thay đổi và thực hiện một số thao tác bổ sung.

### Constants

- **START_SCAN_COL**: Cột bắt đầu quét dữ liệu trong Google Sheets.
- **LEFT_RANGE, LEFT_RANGE_TITLE, LEFT_START_COL, LEFT_END_COL**: Các hằng số xác định phạm vi ô trong Google Sheets.
- **FIELDS_MAPPING**: Mảng ánh xạ các trường dữ liệu từ cơ sở dữ liệu sang Google Sheets.

## Processing stream

### 1. Initialization

- **Constructor**: Nhận một đối tượng `SpreadsheetsService` để tương tác với Google Sheets.

### 2. Method `loadDataToCellFromDatabase`

- **Nhập liệu từ cơ sở dữ liệu**: Tải dữ liệu của một dự án cụ thể vào Google Sheets.

#### Details step by step:

1. **Lấy dự án theo ID**:
   - Kiểm tra xem dự án có tồn tại không.
   - Nếu không tồn tại, ném ra ngoại lệ `Exception`.

2. **Ghi log thông tin dự án**:
   - Ghi lại ID và số dự án vào log.

3. **Đếm số lượng dự án**:
   - Gọi phương thức `getTotalProjectsCount` để đếm tổng số dự án thỏa mãn điều kiện.

4. **Tính toán số trang**:
   - Tính số trang cần thiết dựa trên tổng số dự án.

5. **Khởi tạo biến offset và startRowIndex**:
   - Sử dụng để theo dõi vị trí bắt đầu và vị trí của các dự án trong Google Sheets.

6. **Vòng lặp tải dữ liệu dự án**:
   - Tiếp tục vòng lặp cho đến khi tải xong tất cả các dự án.
   - Trong mỗi lần lặp:
     - Gọi phương thức `fetchProjects` để lấy danh sách dự án theo tiêu chí.
     - Nếu không còn dự án nào, thoát khỏi vòng lặp.
     - Chuẩn bị dữ liệu cho Google Sheets bằng `prepareDataForSheets`.
     - Đảm bảo rằng bảng tính tồn tại với `ensureSheetExists`.
     - Định nghĩa phạm vi cập nhật và gọi `updateGoogleSheet` để cập nhật dữ liệu vào Google Sheets.
     - Ghi log phản hồi cập nhật.
     - Nếu đang cập nhật, ghi log các thay đổi của dự án.

7. **Xử lý các cột động cho sheet lịch**:
   - Gọi `processDynamicColumnsForCalendarSheet` để thực hiện các tác vụ bổ sung cho sheet lịch.

### 3. Secondary method

- **`getTotalProjectsCount`**: Đếm số lượng dự án trong tháng hiện tại, bao gồm cả dự án đã bị xóa.
- **`fetchProjects`**: Lấy danh sách dự án theo tháng hiện tại, với điều kiện giới hạn và sắp xếp.
- **`prepareDataForSheets`**: Chuẩn bị dữ liệu từ các dự án để ghi vào Google Sheets.
- **`ensureSheetExists`**: Kiểm tra và thêm bảng tính mới nếu cần.
- **`getSheetRange`**: Xác định phạm vi ô cho việc cập nhật dữ liệu.
- **`updateGoogleSheet`**: Cập nhật dữ liệu vào Google Sheets.
- **`logUpdateResponse`**: Ghi log phản hồi từ việc cập nhật dữ liệu.
- **`logProjectChanges`**: Ghi log các thay đổi của dự án.
- **`trackProjectChanges`**: So sánh và ghi lại các thay đổi giữa bản ghi hiện tại và dự án đã chọn.
- **`saveProjectHistory`**: Lưu lịch sử thay đổi của dự án vào cơ sở dữ liệu.

## Pseudo code

```plaintext
class ProjectSheetServiceV2:
    constants:
        START_SCAN_COL = 'K'
        LEFT_RANGE = 'A9:J'
        LEFT_RANGE_TITLE = 'A6:K8'
        LEFT_START_COL = 'A'
        LEFT_END_COL = 'J'
        FIELDS_MAPPING = { ... }

    constructor(SpreadsheetsService spread_sheets_service):
        this.spread_sheets_service = spread_sheets_service

    function loadDataToCellFromDatabase(sheetTitle, projectId, isUpdating, currentMonth, currentYear):
        project = fetchProjectById(projectId)
        if not project:
            throw Exception("Project not found")
        
        log("Project ID: " + projectId + ", Project No: " + project.no)
        
        total_project_num = getTotalProjectsCount(currentMonth)
        log("Total Project Num: " + total_project_num)
        
        num = total_project_num
        totalPages = ceil(total_project_num / num)
        log("Total Pages: " + totalPages)

        offset = 0
        startRowIndex = 9

        while offset < total_project_num:
            projects = fetchProjects(currentMonth, offset, num)
            if empty(projects):
                log("No more projects to load.")
                break

            data = prepareDataForSheets(projects)
            ensureSheetExists(sheetTitle)
            range = getSheetRange(sheetTitle, startRowIndex, count(projects))

            response = updateGoogleSheet(range, data)
            logUpdateResponse(response)

            if isUpdating:
                logProjectChanges(sheetTitle, project.no, projects, sheetTitle)

            offset += num
            startRowIndex += num
            processDynamicColumnsForCalendarSheet(sheetTitle, currentMonth, currentYear)



