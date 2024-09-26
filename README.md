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


