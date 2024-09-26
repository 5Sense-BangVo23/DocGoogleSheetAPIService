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
```

# ProjectSheetServiceV2
Class `ProjectSheetServiceV2` chịu trách nhiệm cập nhật dữ liệu từ cơ sở dữ liệu vào Google Sheets cho các dự án. Nó quản lý việc tải dữ liệu, đảm bảo sự tồn tại của bảng, và ghi lại các thay đổi của dự án.

### Constants
- **`START_SCAN_COL`**, **`LEFT_RANGE`**, **`LEFT_RANGE_TITLE`**, **`LEFT_START_COL`**, **`LEFT_END_COL`**: Các hằng số định nghĩa các cột và dải ô trong Google Sheets.
- **`FIELDS_MAPPING`**: Bản đồ các trường từ cơ sở dữ liệu sang tên cột trong Google Sheets.

### Constructor
- Nhận dịch vụ `SpreadsheetsService` để tương tác với Google Sheets.

### LoadDataToCellFromDatabase Method
- **`loadDataToCellFromDatabase($sheetTitle, $projectId, $isUpdating, $currentMonth, $currentYear)`**: Tải dữ liệu từ cơ sở dữ liệu và cập nhật vào Google Sheets.

## Detailed Processing Flow

### Bước 1: Lấy Dự Án
1. **Lấy thông tin dự án** từ cơ sở dữ liệu bằng `projectId`.
   - Nếu không tìm thấy dự án, ném ngoại lệ.

### Bước 2: Đếm Dự Án
2. **Đếm tổng số dự án** theo tiêu chí tháng hiện tại bằng phương thức `getTotalProjectsCount($currentMonth)`.

### Bước 3: Xử Lý Dữ Liệu Theo Trang
3. **Khởi tạo biến** cho việc phân trang: `offset`, `startRowIndex`.
4. **Vòng lặp** để tải dữ liệu theo trang cho đến khi tất cả dự án được xử lý:
   - **Bước 3.1**: Lấy danh sách dự án với phương thức `fetchProjects($currentMonth, $offset, $num)`.
   - Nếu không có dự án nào, ghi log và thoát khỏi vòng lặp.
   - **Bước 3.2**: Chuẩn bị dữ liệu cho Google Sheets bằng `prepareDataForSheets($projects)`.
   - **Bước 3.3**: Đảm bảo bảng tính tồn tại bằng phương thức `ensureSheetExists($sheetTitle)`.
   - **Bước 3.4**: Xác định dải ô cần cập nhật bằng phương thức `getSheetRange($sheetTitle, $startRowIndex, count($projects))`.
   - **Bước 3.5**: Cập nhật dữ liệu vào Google Sheets bằng phương thức `updateGoogleSheet($range, $data)`.
   - **Bước 3.6**: Ghi lại phản hồi thành công từ việc cập nhật dữ liệu.
   - **Bước 3.7**: Nếu `isUpdating` là `true`, ghi lại các thay đổi của dự án bằng `logProjectChanges($sheetTitle, $projectNo, $projects, $sheetTitle)`.
   - **Bước 3.8**: Cập nhật `offset` và `startRowIndex` cho lần lặp tiếp theo.
   - **Bước 3.9**: Gọi `processDynamicColumnsForCalendarSheet($sheetTitle, $currentMonth, $currentYear)` để xử lý các cột động cho bảng lịch.

### Bước 4: Xử Lý Ngoại Lệ
5. Nếu có lỗi xảy ra trong quá trình tải dữ liệu, ghi lại lỗi và ném ngoại lệ.

### Secondary methods
- **`getTotalProjectsCount($currentMonth)`**: Đếm số dự án dựa trên tháng hiện tại và trạng thái.
- **`fetchProjects($currentMonth, $offset, $num)`**: Lấy danh sách dự án từ cơ sở dữ liệu với điều kiện đã chỉ định.
- **`prepareDataForSheets($projects)`**: Chuyển đổi dữ liệu dự án sang định dạng mà Google Sheets yêu cầu.
- **`ensureSheetExists($sheetTitle)`**: Kiểm tra và thêm mới bảng tính nếu chưa tồn tại.
- **`getSheetRange($sheetTitle, $startRowIndex, $projectCount)`**: Xác định dải ô cần cập nhật trong bảng tính.
- **`updateGoogleSheet($range, $data)`**: Gửi yêu cầu cập nhật đến Google Sheets.
- **`logUpdateResponse($response)`**: Ghi lại phản hồi từ việc cập nhật dữ liệu.
- **`logProjectChanges($sheetTitle, $projectNo, $projects, $sheet)`**: Ghi lại các thay đổi của dự án.
- **`trackProjectChanges($currentRecord, $selectedProject, $sheetTitle)`**: So sánh và ghi lại sự thay đổi giữa các trường.
- **`saveProjectHistory($projectId, array $changes, string $sheetTitle)`**: Lưu lại lịch sử thay đổi của dự án vào cơ sở dữ liệu.

## Pseudo code
```
function loadDataToCellFromDatabase(sheetTitle, projectId, isUpdating, currentMonth, currentYear):
    # Lấy thông tin dự án từ cơ sở dữ liệu dựa trên ID dự án
    project = getProjectById(projectId)
    # Kiểm tra xem dự án có tồn tại không
    if project is null:
        throw Exception("Không tìm thấy dự án")

    # Lấy tổng số dự án trong tháng hiện tại
    totalProjects = getTotalProjectsCount(currentMonth)
    num = totalProjects  # Số lượng dự án sẽ được lấy
    offset = 0  # Biến dịch chuyển để theo dõi số lượng dự án đã lấy
    startRowIndex = 9  # Chỉ số hàng bắt đầu trong bảng tính để cập nhật dữ liệu

    # Vòng lặp để tải dữ liệu dự án vào bảng tính
    while offset < totalProjects:
        # Lấy danh sách các dự án dựa trên tháng, offset và số lượng dự án
        projects = fetchProjects(currentMonth, offset, num)
        # Kiểm tra xem có dự án nào được tải về không
        if projects is empty:
            log("Không còn dự án nào để tải")
            break

        # Chuẩn bị dữ liệu cho bảng tính từ danh sách dự án
        data = prepareDataForSheets(projects)
        # Đảm bảo rằng bảng tính đã tồn tại trước khi cập nhật
        ensureSheetExists(sheetTitle)
        # Xác định dải ô trong bảng tính để cập nhật dữ liệu
        range = getSheetRange(sheetTitle, startRowIndex, count(projects))
        # Cập nhật bảng tính Google với dữ liệu đã chuẩn bị
        response = updateGoogleSheet(range, data)
        # Ghi lại phản hồi từ việc cập nhật
        logUpdateResponse(response)

        # Nếu đang cập nhật, ghi lại các thay đổi của dự án
        if isUpdating:
            logProjectChanges(sheetTitle, project.no, projects, sheetTitle)

        # Cập nhật offset và chỉ số hàng bắt đầu cho lần lặp tiếp theo
        offset += num
        startRowIndex += num
        # Xử lý truy cập các cột động cho bảng lịch và project
        processDynamicColumnsForCalendarSheet(sheetTitle, currentMonth, currentYear)

```

# processDynamicColumnsForCalendarSheet method

Hàm `processDynamicColumnsForCalendarSheet` xử lý các cột động cho bảng lịch trong Google Sheets, lấy dữ liệu từ bảng tính và cập nhật các ô tương ứng với các ngày trong tháng hiện tại. Nó cũng đánh dấu các dự án đã bị xóa trên bảng tính.

## Function analysis 

### 1. Params

- **`$sheetTitle`**: Tên của bảng tính trong Google Sheets nơi chứa dữ liệu dự án.
- **`$currentMonth`**: Tháng hiện tại cần xử lý (dưới dạng số, ví dụ: 1 cho tháng 1).
- **`$currentYear`**: Năm hiện tại cần xử lý (dưới dạng số).

### 2. Input

- Sử dụng `spread_sheets_service` để lấy `spreadsheetId`.
- Xác định phạm vi ô cho dự án bằng cách kết hợp tên bảng tính với một hằng số `LEFT_RANGE`.

### 3. Get data từ Google Sheets

- Gọi phương thức `spreadsheets_values->get` của dịch vụ Google Sheets để lấy dữ liệu từ phạm vi đã xác định.
- Lưu trữ kết quả trả về trong biến `$responseProject`.

#### 3.1. Kiểm tra dữ liệu

- Gọi phương thức `getValues()` trên `$responseProject` để lấy giá trị dữ liệu dự án.
- Kiểm tra xem `$valuesProject` có phải là một mảng hay không. Nếu không, ném ra một ngoại lệ với thông báo "No data found in Google Sheets."

### 4. Xác định số ngày trong tháng

- Tạo một đối tượng `DateTime` với ngày đầu tháng hiện tại (ngày 1) để lấy số ngày trong tháng bằng cách gọi phương thức `format('t')`.
- Lưu số ngày này vào biến `$daysInMonth`.

### 5. Tạo hiển lịch theo tháng hiện hành 

- Khởi tạo một mảng `$calendarMap` để ánh xạ ngày tháng với các cột trong bảng tính.
- Sử dụng vòng lặp `for` để duyệt qua từng ngày trong tháng:
  - Sử dụng `chr(ord(self::START_SCAN_COL) + $day - 1)` để tính toán tên cột tương ứng với mỗi ngày (Bắt đầu hiển thị từ column K -> ... Tổng 31 columns tùy theo số ngày trong tháng).
  - Lưu giá trị này vào `$calendarMap` với khóa là số ngày.

### 6. Cập nhật ô cell cho lịch dựa theo thông từ project

- Gọi phương thức `updateCellsForCalendar`, truyền vào các tham số: `$valuesProject`, `$calendarMap`, `$sheetTitle`, và `$currentMonth`.
- Phương thức này sẽ xử lý việc cập nhật các ô trong bảng tính với dữ liệu từ `$valuesProject` dựa trên bản đồ lịch đã tạo.

### 7. Đánh dấu các project đã bị xóa

- Gọi phương thức `highlightDeletedProjectsInSheet`, truyền vào tham số `$sheetTitle`.
- Phương thức này sẽ thực hiện việc đánh dấu các dự án đã bị xóa trong bảng tính.

## Pseudo code

```
function processDynamicColumnsForCalendarSheet(sheetTitle, currentMonth, currentYear):
    # Lấy ID của bảng tính Google Sheets
    spreadsheetId = getSpreadsheetId()

    # Tạo dải ô để lấy dữ liệu từ bảng tính dựa trên tên bảng và phạm vi cố định
    rangeProject = sheetTitle + "!" + LEFT_RANGE
    
    # Lấy giá trị từ Google Sheets trong dải ô đã chỉ định
    responseProject = getGoogleSheetValues(spreadsheetId, rangeProject)
    valuesProject = responseProject.getValues()

    # Kiểm tra xem có dữ liệu không
    if not isArray(valuesProject):
        throw Exception("Không có dữ liệu nào được tìm thấy trong Google Sheets.")
    
    # Lấy số ngày trong tháng hiện tại
    daysInMonth = getDaysInMonth(currentYear, currentMonth)
    
    # Tạo ánh xạ giữa ngày và tên cột trong bảng tính
    calendarMap = {}
    for day from 1 to daysInMonth:
        # Tính toán tên cột tương ứng với ngày
        columnName = chr(ord(START_SCAN_COL) + day - 1)
        calendarMap[day] = columnName

    # Cập nhật các ô trong bảng tính với thông tin từ các dự án
    updateCellsForCalendar(valuesProject, calendarMap, sheetTitle, currentMonth)

    # Làm nổi bật các dự án đã bị xóa trong bảng tính
    highlightDeletedProjectsInSheet(sheetTitle)

```

# updateCellsForCalendar Method

Hàm `updateCellsForCalendar` thực hiện nhiệm vụ cập nhật các ô trong bảng tính Google Sheets với thông tin từ lịch của các dự án. Hàm này xử lý dữ liệu dự án và thực hiện các yêu cầu cập nhật theo từng hàng dữ liệu.

## Endpoint

- Cập nhật các ô tương ứng trong bảng tính với các thông tin ngày quan trọng của các dự án, bao gồm ngày bắt đầu, ngày giao hàng, ngày xuất hàng, và ngày vận chuyển.

## Function analysis

### 1. Params

- **`array $valuesProject`**: Mảng chứa dữ liệu dự án được lấy từ Google Sheets, mỗi hàng tương ứng với một dự án.
- **`array $calendarMap`**: Mảng ánh xạ các ngày trong tháng với các cột trong bảng tính, cho phép xác định vị trí của mỗi ngày trong bảng.
- **`$sheetTitle`**: Tên của bảng tính nơi cần cập nhật dữ liệu.
- **`$currentMonth`**: Tháng hiện tại mà hàm đang xử lý.

### 2. Input

- Sử dụng dịch vụ `spread_sheets_service` để lấy `spreadsheetId`, đại diện cho bảng tính Google Sheets hiện tại.
- Khởi tạo một mảng rỗng `$requests` để chứa các yêu cầu cập nhật mà sẽ được thực hiện sau này.

### 3. Lấy ID bảng tính

- Gọi phương thức `getSheetIdByTitle` để lấy ID của bảng tính dựa trên tên bảng tính được cung cấp (`$sheetTitle`). Điều này giúp xác định bảng mà các ô sẽ được cập nhật.

### 4. Xử lý từng hàng dữ liệu

- Duyệt qua từng hàng dữ liệu trong `$valuesProject` bằng vòng lặp `foreach`:

#### 4.1. Bỏ qua hàng trống

- Kiểm tra xem hàng hiện tại có trống hay không bằng cách sử dụng `empty($row)`. Nếu hàng rỗng, hàm sẽ sử dụng câu lệnh `continue` để bỏ qua và tiếp tục với hàng tiếp theo.

#### 4.2. Kiểm tra projects trên bảng tính 

- Lấy giá trị số dự án từ cột đầu tiên của hàng (`$row[0]`). Nếu giá trị này không tồn tại, biến `$no` sẽ được gán giá trị rỗng (`''`).
- Nếu giá trị `$no` không rỗng, thì thực hiện các bước sau:
  - **Làm sạch giá trị số**: Sử dụng hàm `trim()` để xóa các khoảng trắng ở đầu và cuối chuỗi. 
  - **Kiểm tra kiểu dữ liệu**: Kiểm tra xem giá trị có phải là số không. Nếu đúng, chuyển đổi nó sang kiểu số nguyên.

#### 4.3. Kiểm tra projects từ database

- Sử dụng Eloquent ORM để tìm kiếm dự án trong bảng `cmn_projects` theo số dự án đã lấy. Chỉ lấy các trường cần thiết là `work_start_date`, `dispatch_date`, `shipping_date`, và `delivery_date`.
- Nếu không tìm thấy dự án (`if (!$project)`), ném ra một ngoại lệ với thông báo: `"Project not found for no {$noVal}"`. Ném phát hiện lỗi khi có dự án không tồn tại trong cơ sở dữ liệu.

#### 4.4. Xử lý các ngày

- **Xử lý khoảng thời gian**: Gọi phương thức `processDateRange` để xử lý khoảng thời gian từ `work_start_date` đến `shipping_date`, truyền vào các tham số như:
  - `$requests`: Mảng chứa các yêu cầu cập nhật.
  - `$project->work_start_date`: Ngày bắt đầu của dự án.
  - `$project->shipping_date`: Ngày kết thúc của dự án.
  - `$calendarMap`: Mảng ánh xạ ngày đến cột.
  - `$sheetId`: ID của bảng tính.
  - `$i`: Chỉ số hàng hiện tại trong vòng lặp.
  - `$currentMonth`: Tháng hiện tại.

- **Xử lý ngày xuất hàng**: Gọi phương thức `processDate` với các tham số như sau:
  - `$requests`: Mảng chứa yêu cầu cập nhật.
  - `$project->dispatch_date`: Ngày xuất hàng.
  - `$calendarMap`: Mảng ánh xạ ngày đến cột.
  - `$sheetId`: ID của bảng tính.
  - `$i`: Chỉ số hàng hiện tại.
  - `'発送'`: Nhãn cho ngày xuất hàng.
  - `[1.0, 1.0, 0.0]`: Màu sắc nền cho ô (màu vàng).
  - `[1.0, 1.0, 1.0]`: Màu chữ cho ô (màu trắng).

- **Xử lý ngày vận chuyển**: Tương tự như trên, gọi `processDate` cho ngày `shipping_date` với các tham số tương tự, nhưng sử dụng nhãn `'出荷'` và màu sắc `[0.0, 0.0, 1.0]` cho nền (màu xanh).

### 5. Xử lý ngày giao hàng

- Lấy giá trị ngày giao hàng từ cột thứ 9 của hàng (`$row[8]`).
- Nếu giá trị này không rỗng, gọi phương thức `processDeliveryDate` để xử lý ngày giao hàng với các tham số như:
  - `$requests`: Mảng chứa yêu cầu cập nhật.
  - `$deliveryDateValue`: Giá trị ngày giao hàng.
  - `$calendarMap`: Mảng ánh xạ ngày đến cột.
  - `$sheetId`: ID của bảng tính.
  - `'納品'`: Nhãn cho ngày giao hàng.
  - `$i`: Chỉ số hàng hiện tại.
  - `$currentMonth`: Tháng hiện tại.

### 6. Cập nhật bảng tính theo lô

- Kiểm tra số lượng yêu cầu trong mảng `$requests`. Nếu số lượng yêu cầu vượt quá 100 (`if (count($requests) > 100)`) giá trị phán đoán, thực hiện các bước sau:
  - Gọi phương thức `batchUpdateSpreadsheet` để cập nhật bảng tính với các yêu cầu đã thu thập. Phương thức này giúp gửi nhiều yêu cầu trong một lần để tối ưu hóa hiệu suất.
  - Sau khi cập nhật, làm rỗng mảng `$requests` để chuẩn bị cho các yêu cầu tiếp theo.

### 7. Cập nhật bảng tính lần cuối

- Kiểm tra xem mảng `$requests` có còn yêu cầu nào không. Nếu có (`if (!empty($requests))`), gọi lại `batchUpdateSpreadsheet` để thực hiện cập nhật với các yêu cầu còn lại.
- Nếu không có yêu cầu nào (`else`), ghi thông báo vào log: `"No requests to update spreadsheet."`. Theo dõi các trường hợp không có yêu cầu nào cần cập nhật.

## Pseudo code

```
function updateCellsForCalendar(valuesProject, calendarMap, sheetTitle, currentMonth):
    spreadsheetId = getSpreadsheetId() // Lấy ID của bảng tính
    requests = [] // Khởi tạo mảng yêu cầu cập nhật
    
    sheetId = getSheetIdByTitle(sheetTitle) // Lấy ID của bảng tính theo tiêu đề
    
    for i from 0 to length(valuesProject):
        row = valuesProject[i] // Lấy hàng dữ liệu hiện tại
        
        if isEmpty(row): // Kiểm tra hàng trống
            continue // Bỏ qua hàng trống
        
        no = row[0] or '' // Lấy số dự án
        
        if not isEmpty(no): // Nếu số dự án không rỗng
            noVal = trim(no) // Loại bỏ khoảng trắng ..., làm sạch giá trị 
            noVal = isNumeric(noVal) ? convertToInteger(noVal) : noVal // Kiểm tra kiểu dữ liệu
            
            project = findProjectInDatabase(noVal) // Tìm kiếm dự án với giá trị field:  no
            
            if not project: // Nếu không tìm thấy dự án
                throw Exception("Project not found for no " + noVal) // Ném ngoại lệ
                
            processDateRange(requests, project.work_start_date, project.shipping_date, calendarMap, sheetId, i, currentMonth) // Xử lý khoảng thời gian
            processDate(requests, project.dispatch_date, calendarMap, sheetId, i, '発送', [1.0, 1.0, 0.0], [1.0, 1.0, 1.0]) // Xử lý ngày xuất hàng
            processDate(requests, project.shipping_date, calendarMap, sheetId, i, '出荷', [0.0, 0.0, 1.0], [1.0, 1.0, 1.0]) // Xử lý ngày vận chuyển

        deliveryDateValue = row[8] or '' // Lấy ngày giao hàng
        
        if not isEmpty(deliveryDateValue): // Nếu ngày giao hàng không rỗng
            processDeliveryDate(requests, deliveryDateValue, calendarMap, sheetId, '納品', i, currentMonth) // Xử lý ngày giao hàng
            
    if count(requests) > 100: // Nếu số lượng yêu cầu lớn hơn 100 giá trị phán đoán
        batchUpdateSpreadsheet(spreadsheetId, requests) // Cập nhật bảng tính
        requests = [] // Làm rỗng mảng yêu cầu
        
    if not isEmpty(requests): // Nếu còn yêu cầu nào
        batchUpdateSpreadsheet(spreadsheetId, requests) // Cập nhật bảng tính với các yêu cầu còn lại
    else:
        log("No requests to update spreadsheet.") // Ghi log nếu không có yêu cầu nào

```


# Analyze support functions for `updateCellsForCalendar`

## 1. `processDate(&$requests, $date, $calendarMap, $sheetId, $index, $label, $bgColor, $fgColor)`

### Enpoint
Hàm này xử lý một ngày cụ thể và tạo yêu cầu cập nhật cho ngày đó trong bảng tính.

### Params
- `&$requests`: Tham chiếu đến mảng yêu cầu sẽ được cập nhật.
- `$date`: Ngày cần xử lý (định dạng `YYYY-MM-DD`).
- `$calendarMap`: Bản đồ ánh xạ giữa ngày trong tháng và cột trong bảng tính.
- `$sheetId`: ID của bảng tính cần cập nhật.
- `$index`: Chỉ số hàng cho ngày này.
- `$label`: Nhãn (label) để hiển thị trên ô trong bảng tính.
- `$bgColor`: Màu nền cho ô (mảng RGB).
- `$fgColor`: Màu chữ cho ô (mảng RGB).

### Procedure
1. **Phân tích ngày**: Ngày được phân tích thành năm, tháng và ngày.
2. **Kiểm tra ánh xạ cột**: Nếu có ánh xạ cho ngày, thêm dữ liệu vào mảng `batchData`.
3. **Ghi log cảnh báo**: Nếu không tìm thấy ánh xạ cho ngày, ghi lại cảnh báo.
4. **Thêm yêu cầu**: Nếu có dữ liệu trong `batchData`, sử dụng `addBatchRequests` để thêm yêu cầu vào `$requests`.

---

## 2. `processDateRange(&$requests, $startDate, $endDate, $calendarMap, $sheetId, $index, $currentMonth)`

### Endpoint
Hàm này xử lý một khoảng thời gian từ ngày bắt đầu đến ngày kết thúc và tạo yêu cầu cập nhật cho từng ngày trong khoảng thời gian đó.

### Params
- `&$requests`: Tham chiếu đến mảng yêu cầu sẽ được cập nhật.
- `$startDate`: Ngày bắt đầu (định dạng `YYYY-MM-DD`).
- `$endDate`: Ngày kết thúc (định dạng `YYYY-MM-DD`).
- `$calendarMap`: Bản đồ ánh xạ giữa ngày trong tháng và cột trong bảng tính.
- `$sheetId`: ID của bảng tính cần cập nhật.
- `$index`: Chỉ số hàng cho ngày trong khoảng thời gian.
- `$currentMonth`: Tháng hiện tại để kiểm tra.

### Procedure
1. **Kiểm tra ngày**: Nếu cả hai ngày bắt đầu và kết thúc đều không rỗng, tiếp tục xử lý.
2. **Lặp qua các ngày**: Sử dụng vòng lặp để đi qua từng ngày trong khoảng thời gian, kiểm tra xem ngày đó có thuộc tháng hiện tại và có ánh xạ cột không.
3. **Thêm dữ liệu vào batchData**: Nếu thỏa mãn điều kiện, thêm dữ liệu vào mảng `batchData`.
4. **Thêm yêu cầu**: Nếu có dữ liệu trong `batchData`, sử dụng `addBatchRequests` để thêm yêu cầu vào `$requests`.

---

## 3. `processDeliveryDate(array &$requests, $deliveryDateValue, array $calendarMap, $sheetId, $label, $index, $currentMonth)`

### Endpoint
Hàm này xử lý ngày giao hàng và tạo yêu cầu cập nhật cho ngày đó trong bảng tính.

### Params
- `array &$requests`: Tham chiếu đến mảng yêu cầu sẽ được cập nhật.
- `$deliveryDateValue`: Giá trị ngày giao hàng cần xử lý.
- `array $calendarMap`: Bản đồ ánh xạ giữa ngày trong tháng và cột trong bảng tính.
- `$sheetId`: ID của bảng tính cần cập nhật.
- `$label`: Nhãn (label) để hiển thị trên ô trong bảng tính.
- `$index`: Chỉ số hàng cho ngày giao hàng.
- `$currentMonth`: Tháng hiện tại để kiểm tra.

### Procedure
1. **Phân tích ngày giao hàng**: Sử dụng hàm `parseDeliveryDate` để lấy tháng và ngày từ giá trị ngày giao hàng.
2. **Kiểm tra ánh xạ cột**: Nếu tháng của ngày giao hàng trùng với tháng hiện tại và có ánh xạ cho ngày đó, thêm dữ liệu vào mảng `batchData`.
3. **Thêm yêu cầu**: Nếu có dữ liệu trong `batchData`, sử dụng `addBatchRequests` để thêm yêu cầu vào `$requests`.

## 4. `addBatchRequests(&$requests, $sheetColumn, $sheetId, $data)`

### Endpoint
Hàm này thêm các yêu cầu cập nhật cho các ô trong bảng tính vào mảng `$requests`.

### Params
- `&$requests`: Tham chiếu đến mảng yêu cầu sẽ được cập nhật.
- `$sheetColumn`: Cột trong bảng tính nơi ô sẽ được cập nhật.
- `$sheetId`: ID của bảng tính cần cập nhật.
- `$data`: Dữ liệu chứa thông tin về ngày, nhãn, màu nền và màu chữ.

### Procedure
1. **Chuẩn bị hàng**: Lặp qua từng mục trong `data` và chuẩn bị hàng cho yêu cầu cập nhật.
2. **Xác định chỉ số hàng và cột**: Tính toán chỉ số hàng bắt đầu và kết thúc cũng như chỉ số cột để cập nhật.
3. **Thêm yêu cầu vào mảng**: Thêm yêu cầu cập nhật vào mảng `$requests` với thông tin định dạng và giá trị ô.

---

## 5. `batchUpdateSpreadsheet($spreadsheetId, array $requests, $type)`

### Endpoint
Hàm này gửi yêu cầu cập nhật đến bảng tính Google Sheets.

### Params
- `$spreadsheetId`: ID của bảng tính cần cập nhật.
- `array $requests`: Mảng các yêu cầu cập nhật.
- `$type`: Loại yêu cầu cập nhật (được sử dụng cho ghi log).

### Procedure
1. **Kiểm tra yêu cầu**: Nếu không có yêu cầu nào trong `$requests`, ghi log thông báo rằng không có ô nào để cập nhật.
2. **Gửi yêu cầu cập nhật**: Tạo một đối tượng `BatchUpdateSpreadsheetRequest` và gửi yêu cầu bằng cách sử dụng `google_sheet_service`.
3. **Ghi log lỗi**: Nếu có lỗi xảy ra trong quá trình gửi yêu cầu, ghi lại thông báo lỗi.

---

## 6. `parseDeliveryDate($deliveryDate)`

### Endpoint
Hàm này phân tích giá trị ngày giao hàng và trả về tháng và ngày.

### Params
- `$deliveryDate`: Giá trị ngày giao hàng (định dạng `MM月DD日`).

### Procedure
1. **Sử dụng biểu thức chính quy**: Áp dụng biểu thức chính quy để tìm tháng và ngày từ giá trị đầu vào.
2. **Kiểm tra và trả về**: Nếu tìm thấy tháng và ngày, trả về chúng dưới dạng mảng. Nếu không, trả về `[null, null]`.

---

## 7. `getFieldColumnIndices(array $header, array $fieldMapping): array`

### Endpoint
Hàm này lấy chỉ số cột tương ứng từ tiêu đề của bảng tính dựa trên ánh xạ trường.

### Params
- `array $header`: Mảng tiêu đề của bảng tính.
- `array $fieldMapping`: Mảng ánh xạ giữa tiêu đề và trường.

### Procedure
1. **Tạo bản đồ chỉ số tiêu đề**: Sử dụng `array_flip` để tạo bản đồ ánh xạ từ tiêu đề sang chỉ số.
2. **Lặp qua ánh xạ trường**: Kiểm tra từng trường trong ánh xạ và thêm chỉ số cột vào `fieldColumnIndex`.
3. **Ghi log cảnh báo**: Nếu tiêu đề không tìm thấy trong tiêu đề của bảng tính, ghi lại cảnh báo.
4. **Trả về chỉ số cột**: Trả về mảng chỉ số cột tương ứng với các trường.







