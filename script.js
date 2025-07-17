class ExcelAnalyzer {
  constructor() {
    this.file = null
    this.isProcessing = false
    this.initializeEventListeners()
  }

  initializeEventListeners() {
    const fileInput = document.getElementById("file-upload")
    const analyzeBtn = document.getElementById("analyze-btn")

    fileInput.addEventListener("change", (e) => this.handleFileUpload(e))
    analyzeBtn.addEventListener("click", () => this.analyzeExcel())
  }

  handleFileUpload(event) {
    const file = event.target.files[0]
    const fileStatus = document.getElementById("file-status")
    const analyzeBtn = document.getElementById("analyze-btn")

    if (file) {
      // Check if file is Excel format
      const validTypes = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
        ".xlsx",
        ".xls",
      ]

      const isValidFile = validTypes.some((type) => file.type === type || file.name.toLowerCase().endsWith(type))

      if (isValidFile) {
        this.file = file
        fileStatus.textContent = `✅ Selected: ${file.name}`
        fileStatus.className = "file-status success"
        analyzeBtn.disabled = false
        this.hideMessages()
      } else {
        fileStatus.textContent = "❌ Please select a valid Excel file (.xlsx or .xls)"
        fileStatus.className = "file-status error"
        analyzeBtn.disabled = true
        this.file = null
      }
    } else {
      this.file = null
      fileStatus.textContent = ""
      fileStatus.className = "file-status"
      analyzeBtn.disabled = true
    }
  }

  hideMessages() {
    document.getElementById("success-message").classList.add("hidden")
    document.getElementById("error-message").classList.add("hidden")
  }

  showError(message) {
    const errorMessage = document.getElementById("error-message")
    const errorText = document.getElementById("error-text")
    errorText.textContent = message
    errorMessage.classList.remove("hidden")
  }

  showSuccess() {
    document.getElementById("success-message").classList.remove("hidden")
  }

  setProcessing(processing) {
    this.isProcessing = processing
    const btn = document.getElementById("analyze-btn")
    const btnText = document.getElementById("btn-text")
    const spinner = document.getElementById("loading-spinner")

    if (processing) {
      btn.disabled = true
      btnText.textContent = "Processing..."
      spinner.classList.remove("hidden")
    } else {
      btn.disabled = !this.file
      btnText.textContent = "📊 Analyze Excel File"
      spinner.classList.add("hidden")
    }
  }

  async analyzeExcel() {
    if (!this.file || this.isProcessing) return

    // Check if XLSX library is loaded
    const XLSX = window.XLSX // Declare the XLSX variable here
    if (typeof XLSX === "undefined") {
      this.showError("Excel processing library not loaded. Please refresh the page and try again.")
      return
    }

    this.setProcessing(true)
    this.hideMessages()

    try {
      const initialUslNo = Number.parseInt(document.getElementById("initial-usl").value)
      const arrayBuffer = await this.file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: "array" })

      // Get worksheets
      const sheet1 = workbook.Sheets["Sheet1"]
      const sheet2 = workbook.Sheets["Sheet2"]

      if (!sheet1 || !sheet2) {
        throw new Error("Required sheets (Sheet1, Sheet2) not found in the Excel file")
      }

      // Get subject codes and names
      const subjectCodeAndName = this.getSubjectCodes(sheet1)

      if (Object.keys(subjectCodeAndName).length === 0) {
        throw new Error("No subject codes found in Sheet1. Please check the format.")
      }

      // Get student data
      const students = this.getStudentData(sheet2, subjectCodeAndName, initialUslNo)

      if (students.length === 0) {
        throw new Error(`No students found starting from USL number ${initialUslNo}`)
      }

      // Calculate statistics
      const stats = this.calculateStatistics(students, subjectCodeAndName)

      // Generate report
      const reportWorkbook = this.generateReport(students, subjectCodeAndName, stats)

      // Download file
      this.downloadFile(reportWorkbook, `Analysis_${this.file.name}`)

      this.showSuccess()
    } catch (error) {
      console.error("Analysis error:", error)
      this.showError(error.message)
    } finally {
      this.setProcessing(false)
    }
  }

  findInWorksheet(ws, searchValue) {
    const range = window.XLSX.utils.decode_range(ws["!ref"] || "A1:A1")

    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = window.XLSX.utils.encode_cell({ r: row, c: col })
        const cell = ws[cellAddress]
        if (cell && String(cell.v).trim() === String(searchValue).trim()) {
          return [row + 1, col + 1] // Convert to 1-based indexing
        }
      }
    }
    return [-1, -1]
  }

  getCellValue(ws, row, col) {
    const cellAddress = window.XLSX.utils.encode_cell({ r: row - 1, c: col - 1 })
    const cell = ws[cellAddress]
    return cell ? cell.v : null
  }

  getSubjectCodes(sheet1) {
    const [courseRow, courseCol] = this.findInWorksheet(sheet1, "Course Code")
    const subjectCodeAndName = {}

    if (courseRow === -1) {
      throw new Error("'Course Code' header not found in Sheet1. Please check the format.")
    }

    let currentRow = courseRow + 1
    let attempts = 0
    const maxAttempts = 100 // Prevent infinite loop

    while (attempts < maxAttempts) {
      const code = this.getCellValue(sheet1, currentRow, courseCol)
      const name = this.getCellValue(sheet1, currentRow, courseCol + 1)

      if (!code) break // No more data

      if (code && name) {
        subjectCodeAndName[String(code).trim()] = String(name).trim()
      }

      currentRow++
      attempts++
    }

    return subjectCodeAndName
  }

  getStudentData(sheet2, subjectCodeAndName, initialUslNo) {
    const initialUslStr = String(initialUslNo).padStart(5, "0")
    const [uslInitRow, uslCol] = this.findInWorksheet(sheet2, initialUslStr)

    if (uslInitRow === -1) {
      throw new Error(`Initial USL number ${initialUslStr} not found in Sheet2`)
    }

    // Get all USL IDs
    const range = window.XLSX.utils.decode_range(sheet2["!ref"] || "A1:A1")
    const uslIds = []

    for (let row = uslInitRow; row <= range.e.r + 1; row++) {
      const value = this.getCellValue(sheet2, row, uslCol)
      if (value && String(value).match(/^\d{5}$/)) {
        uslIds.push(String(value))
      } else if (value === "USN") {
        break
      }
    }

    // Process each student
    const students = []

    for (const uslId of uslIds) {
      const [refRow] = this.findInWorksheet(sheet2, uslId)
      if (refRow === -1) continue

      const student = {
        refNo: refRow,
        uslNo: uslId,
        name: this.getCellValue(sheet2, refRow + 2, uslCol) || "",
        usn: this.getCellValue(sheet2, refRow + 1, uslCol + 1) || "",
        sgpa: this.getCellValue(sheet2, refRow + 5, uslCol + 1) || 0,
        cgpa: this.getCellValue(sheet2, refRow + 6, uslCol + 1) || 0,
        result: this.getResultStatus(this.getCellValue(sheet2, refRow + 7, uslCol) || ""),
        termGrade: this.getTermGrade(this.getCellValue(sheet2, refRow + 9, uslCol) || ""),
        subjects: {},
        total: 0,
      }

      // Get subject marks
      let tempCol = uslCol + 3
      let attempts = 0
      const maxCols = 50 // Prevent infinite loop

      while (attempts < maxCols) {
        const cellValue = this.getCellValue(sheet2, refRow, tempCol)
        if (cellValue === "Total") break

        const subject = cellValue
        if (subject && subjectCodeAndName[subject]) {
          const marks1 = this.getCellValue(sheet2, refRow + 7, tempCol) || 0
          const marks2 = this.getCellValue(sheet2, refRow + 10, tempCol) || ""
          student.subjects[subject] = [Number(marks1), String(marks2)]
        }

        tempCol++
        attempts++
      }

      student.total = this.getCellValue(sheet2, refRow + 7, tempCol) || 0
      students.push(student)
    }

    return students
  }

  getResultStatus(text) {
    if (!text) return ""
    const match = String(text).match(/Result: ([A-Za-z]+)/)
    return match ? match[1] : ""
  }

  getTermGrade(text) {
    if (!text) return ""
    const match = String(text).match(/Term Grade: ([^\n]+)/)
    return match ? match[1].trim() : ""
  }

  calculateStatistics(students, subjectCodeAndName) {
    const subjectStats = {}

    // Initialize subject stats
    Object.keys(subjectCodeAndName).forEach((code) => {
      subjectStats[code] = {
        appeared: 0,
        pass: 0,
        fail: 0,
        percentage: 0,
        average: 0,
        max: 0,
        min: Number.POSITIVE_INFINITY,
      }
    })

    // Calculate stats for each subject
    students.forEach((student) => {
      Object.entries(student.subjects).forEach(([code, [marks, result]]) => {
        if (subjectStats[code]) {
          subjectStats[code].appeared++
          if (result.toLowerCase() === "pass") {
            subjectStats[code].pass++
          } else if (result.toLowerCase() === "fail") {
            subjectStats[code].fail++
          }

          subjectStats[code].max = Math.max(subjectStats[code].max, marks)
          subjectStats[code].min = Math.min(subjectStats[code].min, marks)
        }
      })
    })

    // Calculate percentages and averages
    Object.keys(subjectStats).forEach((code) => {
      const stat = subjectStats[code]
      stat.percentage = stat.appeared > 0 ? (stat.pass / stat.appeared) * 100 : 0

      const subjectMarks = students.map((s) => s.subjects[code]?.[0] || 0).filter((mark) => mark > 0)

      stat.average =
        subjectMarks.length > 0 ? subjectMarks.reduce((sum, mark) => sum + mark, 0) / subjectMarks.length : 0

      if (stat.min === Number.POSITIVE_INFINITY) stat.min = 0
    })

    return subjectStats
  }

  generateReport(students, subjectCodeAndName, stats) {
    const wb = window.XLSX.utils.book_new()

    // Create Final sheet
    const finalData = this.createFinalSheetData(students, subjectCodeAndName)
    const finalWs = window.XLSX.utils.aoa_to_sheet(finalData)
    window.XLSX.utils.book_append_sheet(wb, finalWs, "Final")

    // Create Report sheet
    const reportData = this.createReportSheetData(students, subjectCodeAndName, stats)
    const reportWs = window.XLSX.utils.aoa_to_sheet(reportData)
    window.XLSX.utils.book_append_sheet(wb, reportWs, "Report")

    return wb
  }

  createFinalSheetData(students, subjectCodeAndName) {
    const finalData = []
    const headers = [
      "Ref No",
      "USL no",
      "Name",
      "USN",
      "SGPA",
      "CGPA",
      "Result",
      "Term Grade",
      ...Object.values(subjectCodeAndName),
      "Total",
    ]
    finalData.push(headers)

    students.forEach((student) => {
      const row = [
        student.refNo,
        student.uslNo,
        student.name,
        student.usn,
        student.sgpa,
        student.cgpa,
        student.result,
        student.termGrade,
      ]

      Object.keys(subjectCodeAndName).forEach((code) => {
        row.push(student.subjects[code]?.[0] || 0)
      })

      row.push(student.total)
      finalData.push(row)
    })

    // Add summary statistics
    finalData.push([])
    finalData.push([])

    const totalAppeared = students.filter((s) => ["pass", "fail"].includes(s.result.toLowerCase())).length
    const totalPass = students.filter((s) => s.result.toLowerCase() === "pass").length
    const totalFail = students.filter((s) => s.result.toLowerCase() === "fail").length

    finalData.push(["", "", "Number of students appeared", "", "", totalAppeared])
    finalData.push(["", "", "PASS", "", "", totalPass])
    finalData.push(["", "", "FAIL", "", "", totalFail])
    finalData.push([
      "",
      "",
      "Pass Percentage",
      "",
      "",
      totalAppeared > 0 ? ((totalPass / totalAppeared) * 100).toFixed(2) : 0,
    ])

    return finalData
  }

  createReportSheetData(students, subjectCodeAndName, stats) {
    const reportData = []
    reportData.push([
      "Sl.No",
      "Subject",
      "Faculty Handled",
      "Number of students appeared",
      "PASS",
      "FAIL",
      "Subject Wise %",
      "AVG",
      "MAX",
      "MIN",
    ])

    let slNo = 1
    Object.entries(subjectCodeAndName).forEach(([code, name]) => {
      const stat = stats[code]
      reportData.push([
        slNo++,
        name,
        "",
        stat.appeared,
        stat.pass,
        stat.fail,
        stat.percentage.toFixed(2),
        stat.average.toFixed(2),
        stat.max,
        stat.min,
      ])
    })

    // Add overall statistics
    reportData.push([])
    reportData.push([])

    const termGrades = students.map((s) => s.termGrade)
    const totalAppeared = students.filter((s) => ["pass", "fail"].includes(s.result.toLowerCase())).length
    const totalPass = students.filter((s) => s.result.toLowerCase() === "pass").length
    const totalFail = students.filter((s) => s.result.toLowerCase() === "fail").length

    const gradeStats = [
      ["Total No. of Outstanding - O", termGrades.filter((g) => g === "O").length],
      ["Total No. of A+", termGrades.filter((g) => g === "A+").length],
      ["Total No. of A", termGrades.filter((g) => g === "A").length],
      ["Total No. of B+", termGrades.filter((g) => g === "B+").length],
      ["Total No. of Fail", totalFail],
      ["Total No. of Appeared", totalAppeared],
      ["Total No. of Absentees", students.length - totalAppeared],
      ["Total No. of Students", students.length],
      ["Total Passing Percentage", totalAppeared > 0 ? ((totalPass / totalAppeared) * 100).toFixed(2) : 0],
    ]

    reportData.push(["", "Particulars", "Total"])
    gradeStats.forEach(([particular, total]) => {
      reportData.push(["", particular, total])
    })

    return reportData
  }

  downloadFile(workbook, filename) {
    const wbout = window.XLSX.write(workbook, { bookType: "xlsx", type: "array" })
    const blob = new Blob([wbout], { type: "application/octet-stream" })
    const url = URL.createObjectURL(blob)
    const a = document.createElement("a")
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }
}

// Initialize the application when DOM is loaded
document.addEventListener("DOMContentLoaded", () => {
  // Check if XLSX library is loaded
  const XLSX = window.XLSX // Declare the XLSX variable here
  if (typeof XLSX === "undefined") {
    console.error("XLSX library not loaded")
    document.getElementById("analyze-btn").disabled = true
    document.getElementById("error-message").classList.remove("hidden")
    document.getElementById("error-text").textContent =
      "Excel processing library failed to load. Please refresh the page."
    return
  }

  new ExcelAnalyzer()
})

// Also initialize if DOM is already loaded
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => new ExcelAnalyzer())
} else {
  new ExcelAnalyzer()
}
