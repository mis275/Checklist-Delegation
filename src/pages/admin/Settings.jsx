import React, { useState, useEffect, useMemo } from 'react'
import { useNavigate } from 'react-router-dom'
import AdminLayout from '../../components/layout/AdminLayout'
import { Plus, User, Building, X, Save, Edit, Trash2, Search, ChevronDown, Loader2, Building2, Users, RefreshCw, Pencil, LogOut, Upload, History, Calendar } from 'lucide-react'

// Configuration
const CONFIG = {
    APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbxG7zW6AabjyxnEDh9JIKMp978w_ik7xzcDy1rCygg3UFFDxYZW6D6rAuxcVHRVaE0O/exec",
    SHEET_NAME: "Whatsapp",
    LEAVE_SHEET_NAME: "Unique",
    CHECKLIST_SHEET_NAME: "Checklist"
}

function Settings() {
    const navigate = useNavigate()
    const [data, setData] = useState([])
    const [loading, setLoading] = useState(true)
    const [error, setError] = useState(null)
    const [editingId, setEditingId] = useState(null)
    const [editedValues, setEditedValues] = useState({ doerName: '', password: '', role: '', idEmail: '', number: '' })
    const [transferModalLoading, setTransferModalLoading] = useState(false)
    const [saving, setSaving] = useState(false)
    const [successMessage, setSuccessMessage] = useState('')
    const [leaveNameFilter, setLeaveNameFilter] = useState('')
    const [deleting, setDeleting] = useState(null)
    const [authorized, setAuthorized] = useState(false)
    const [showAddModal, setShowAddModal] = useState(false)
    const [newEntry, setNewEntry] = useState({ doerName: '', password: '', role: '', idEmail: '', number: '' })
    const [adding, setAdding] = useState(false)

    // UI states from reference
    const [activeTab, setActiveTab] = useState('users') // 'users' or 'departments'
    const [usernameFilter, setUsernameFilter] = useState('')
    const [usernameDropdownOpen, setUsernameDropdownOpen] = useState(false)
    const [activeDeptSubTab, setActiveDeptSubTab] = useState('departments') // 'departments' or 'givenBy'

    // Department states
    const [showDeptModal, setShowDeptModal] = useState(false)
    const [newDeptEntry, setNewDeptEntry] = useState({ department: '', givenBy: '' })
    const [addingDept, setAddingDept] = useState(false)
    const [editingDeptId, setEditingDeptId] = useState(null)
    const [editedDeptValues, setEditedDeptValues] = useState({ department: '', givenBy: '' })
    const [savingDept, setSavingDept] = useState(false)

    // Leave states
    const [leaveData, setLeaveData] = useState([])
    const [leaveLoading, setLeaveLoading] = useState(false)
    const [leaveError, setLeaveError] = useState(null)
    const [leaveRemarks, setLeaveRemarks] = useState({}) // { [_rowIndex]: remarksValue }
    const [savingRemarksId, setSavingRemarksId] = useState(null)

    // Transfer/Leave Modal States
    const [showTransferModal, setShowTransferModal] = useState(false)
    const [transferTask, setTransferTask] = useState(null)
    const [transferForm, setTransferForm] = useState({
        startDate: new Date().toISOString().split('T')[0],
        endDate: new Date().toISOString().split('T')[0],
        remarks: '',
        isChecked: false
    })
    const [transferring, setTransferring] = useState(false)
    const [checklistTasks, setChecklistTasks] = useState([])
    const [checklistLoading, setChecklistLoading] = useState(false)
    const [selectedChecklistTaskIds, setSelectedChecklistTaskIds] = useState([])
    const [allChecklistRows, setAllChecklistRows] = useState(null)
    const [checklistHeaders, setChecklistHeaders] = useState([])
    const [debugInfo, setDebugInfo] = useState(null)
    
    // Leave History states
    const [leaveHistoryData, setLeaveHistoryData] = useState([])
    const [leaveHistoryLoading, setLeaveHistoryLoading] = useState(false)
    const [leaveHistoryError, setLeaveHistoryError] = useState(null)
    const [leaveHistoryNameFilter, setLeaveHistoryNameFilter] = useState('')
    const [leaveHistoryStartDate, setLeaveHistoryStartDate] = useState('')
    const [leaveHistoryEndDate, setLeaveHistoryEndDate] = useState('')

    // Memoized unique rows for the leave tab
    const uniqueLeaveRows = useMemo(() => {
        if (!leaveData || leaveData.length === 0) return []
        const headers = leaveData[0]?.headers || []
        const nameIdx = (() => {
            const i = headers.findIndex(h => {
                const l = h?.toLowerCase() || ''
                return l.includes('name') || l.includes('doer') || l.includes('user')
            })
            return i !== -1 ? i : 3
        })()

        const seen = new Set()
        const unique = []
        for (const row of leaveData) {
            const name = String(row.values[nameIdx] || '').trim().toLowerCase()
            if (name && !seen.has(name)) {
                seen.add(name)
                unique.push(row)
            }
        }
        return unique
    }, [leaveData])

    // Fetch leave data when tab is active
    useEffect(() => {
        if (authorized && activeTab === 'leave' && leaveData.length === 0) {
            fetchLeaveData()
        }
    }, [activeTab, authorized, leaveData.length])

    // Fetch leave history when tab is active
    useEffect(() => {
        if (authorized && activeTab === 'leave_history' && leaveHistoryData.length === 0) {
            fetchLeaveHistoryData()
        }
    }, [activeTab, authorized, leaveHistoryData.length])

    // Check authorization on mount
    useEffect(() => {
        const userRole = sessionStorage.getItem('role') || ''
        const normalizedRole = userRole.toLowerCase().trim().replace(/\s+/g, '_')

        // Only allow admin, super_admin or superadmin
        if (normalizedRole === 'admin' || normalizedRole === 'super_admin' || normalizedRole === 'superadmin') {
            setAuthorized(true)
            fetchWhatsappData()
        } else {
            navigate('/dashboard/admin', { replace: true })
        }
    }, [navigate])

    const fetchWhatsappData = async () => {
        try {
            setLoading(true)
            setError(null)

            const response = await fetch(
                `${CONFIG.APPS_SCRIPT_URL}?sheet=${CONFIG.SHEET_NAME}&action=fetch`
            )

            if (!response.ok) {
                throw new Error(`Failed to fetch data: ${response.status}`)
            }

            const responseText = await response.text()
            let parsedData

            try {
                parsedData = JSON.parse(responseText)
            } catch (parseError) {
                const jsonStart = responseText.indexOf("{")
                const jsonEnd = responseText.lastIndexOf("}")
                if (jsonStart !== -1 && jsonEnd !== -1) {
                    const jsonString = responseText.substring(jsonStart, jsonEnd + 1)
                    parsedData = JSON.parse(jsonString)
                } else {
                    throw new Error("Invalid JSON response from server")
                }
            }

            let rows = []
            if (parsedData.table && parsedData.table.rows) {
                rows = parsedData.table.rows
            } else if (Array.isArray(parsedData)) {
                rows = parsedData
            } else if (parsedData.values) {
                rows = parsedData.values.map((row) => ({
                    c: row.map((val) => ({ v: val })),
                }))
            }

            const extractedData = rows
                .slice(1) // Skip header row
                .map((row, index) => {
                    let rowValues = []
                    if (row.c) {
                        rowValues = row.c.map((cell) =>
                            cell && cell.v !== undefined ? cell.v : ""
                        )
                    } else if (Array.isArray(row)) {
                        rowValues = row
                    }

                    // Column A=index 0, B=index 1, D=index 3, E=index 4, F=index 5, G=index 6, H=index 7
                    // Note: Column C (index 2) is used in reference as username, but my map has D as doerName
                    // I will stick to my 8-column map as per the previous working version
                    const department = rowValues[0] || ""
                    const givenBy = rowValues[1] || ""
                    const doerName = rowValues[3] || ""
                    const password = rowValues[4] || ""
                    const role = rowValues[5] || ""
                    const idEmail = rowValues[6] || ""
                    const number = rowValues[7] || ""

                    if (doerName || password || department || givenBy) {
                        return {
                            id: index,
                            _rowIndex: index + 2,
                            department: department.toString().trim(),
                            givenBy: givenBy.toString().trim(),
                            doerName: doerName.toString().trim(),
                            password: password.toString().trim(),
                            role: role.toString().trim(),
                            idEmail: idEmail.toString().trim(),
                            number: number.toString().trim(),
                            _originalRow: rowValues
                        }
                    }
                    return null
                })
                .filter(Boolean)

            setData(extractedData)
        } catch (err) {
            console.error("Error fetching Whatsapp data:", err)
            setError(err.message)
        } finally {
            setLoading(false)
        }
    }

    const fetchLeaveData = async () => {
        try {
            setLeaveLoading(true)
            setLeaveError(null)

            const response = await fetch(
                `${CONFIG.APPS_SCRIPT_URL}?sheet=${CONFIG.LEAVE_SHEET_NAME}&action=fetch`
            )

            if (!response.ok) {
                throw new Error(`Failed to fetch data: ${response.status}`)
            }

            const responseText = await response.text()
            let parsedData

            try {
                parsedData = JSON.parse(responseText)
            } catch (parseError) {
                const jsonStart = responseText.indexOf("{")
                const jsonEnd = responseText.lastIndexOf("}")
                if (jsonStart !== -1 && jsonEnd !== -1) {
                    const jsonString = responseText.substring(jsonStart, jsonEnd + 1)
                    parsedData = JSON.parse(jsonString)
                } else {
                    throw new Error("Invalid JSON response from server")
                }
            }

            let rows = []
            if (parsedData.table && parsedData.table.rows) {
                rows = parsedData.table.rows
            } else if (Array.isArray(parsedData)) {
                rows = parsedData
            } else if (parsedData.values) {
                rows = parsedData.values.map((row) => ({
                    c: row.map((val) => ({ v: val })),
                }))
            }

            const headerRow = rows[0]?.c?.map(cell => cell?.v) || rows[0] || []

            const extractedData = rows
                .slice(2) // Skip header row
                .map((row, index) => {
                    let rowValues = []
                    if (row.c) {
                        rowValues = row.c.map((cell) =>
                            cell && cell.v !== undefined ? cell.v : ""
                        )
                    } else if (Array.isArray(row)) {
                        rowValues = row
                    }

                    return {
                        id: index,
                        _rowIndex: index + 2,
                        values: rowValues,
                        headers: headerRow
                    }
                })
                .filter(Boolean)
            console.log("extract", extractedData)
            setLeaveData(extractedData)
        } catch (err) {
            console.error("Error fetching Leave data:", err)
            setLeaveError(err.message)
        } finally {
            setLeaveLoading(false)
        }
    }

    const fetchLeaveHistoryData = async () => {
        try {
            setLeaveHistoryLoading(true)
            setLeaveHistoryError(null)

            const response = await fetch(
                `${CONFIG.APPS_SCRIPT_URL}?sheet=${CONFIG.CHECKLIST_SHEET_NAME}&action=fetch`
            )

            if (!response.ok) {
                throw new Error(`Failed to fetch data: ${response.status}`)
            }

            const responseText = await response.text()
            let allRows = []

            try {
                const parsed = JSON.parse(responseText)
                if (parsed.table && parsed.table.rows) {
                    allRows = parsed.table.rows
                } else if (Array.isArray(parsed)) {
                    allRows = parsed
                } else if (parsed.values) {
                    allRows = parsed.values.map((row) => ({
                        c: row.map((val) => ({ v: val })),
                    }))
                }
            } catch (parseError) {
                const jsonStart = responseText.indexOf("{")
                const jsonEnd = responseText.lastIndexOf("}")
                if (jsonStart !== -1 && jsonEnd !== -1) {
                    const jsonString = responseText.substring(jsonStart, jsonEnd + 1)
                    const parsed = JSON.parse(jsonString)
                    if (parsed.table && parsed.table.rows) {
                        allRows = parsed.table.rows
                    }
                } else {
                    throw new Error("Invalid JSON response from server")
                }
            }

            if (allRows.length === 0) {
                setLeaveHistoryData([])
                return
            }

            // Get headers from first row
            const headers = allRows[0]?.c?.map(c => c?.v) || allRows[0] || []
            
            // Find column indices
            const remarksIdx = headers.findIndex(h => h?.toString().toLowerCase().includes('remark'))
            const nameIdx = headers.findIndex(h => 
                h?.toString().toLowerCase().includes('name') || 
                h?.toString().toLowerCase().includes('doer') || 
                h?.toString().toLowerCase().includes('user')
            )
            const dateIdx = headers.findIndex(h => 
                h?.toString().toLowerCase().includes('date') || 
                h?.toString().toLowerCase().includes('time')
            )
            // Task ID is in Column F (Index 5)
            const taskIdIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l === 'task id' || l === 'id' || l === 'taskid'
                })
                return i !== -1 ? i : 5
            })()
            // Task Description is in Column B (Index 1)
            const descIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l.includes('description') || l.includes('desc') || l === 'task'
                })
                return i !== -1 ? i : 1
            })()

            const leaves = allRows.slice(1).map((row, idx) => {
                let vals = []
                if (row.c) vals = row.c.map(c => c?.v)
                else if (Array.isArray(row)) vals = row

                const remarkValue = String(vals[remarksIdx] || '').trim()
                const isLeave = remarkValue.toLowerCase().includes('leave')

                if (isLeave) {
                    return {
                        id: idx,
                        _rowIndex: idx + 2,
                        date: vals[dateIdx],
                        name: vals[nameIdx],
                        taskId: vals[taskIdIdx],
                        description: vals[descIdx],
                        remarks: remarkValue,
                        allValues: vals
                    }
                }
                return null
            }).filter(Boolean)

            setLeaveHistoryData(leaves)
        } catch (err) {
            console.error("Error fetching Leave History:", err)
            setLeaveHistoryError(err.message)
        } finally {
            setLeaveHistoryLoading(false)
        }
    }

    // Fetch checklist tasks for a specific user
    const fetchUserChecklistTasks = async (userName) => {
        try {
            setChecklistLoading(true)
            setChecklistTasks([])

            let allRows = []
            let headers = checklistHeaders

            if (allChecklistRows) {
                allRows = allChecklistRows
            } else {
                const response = await fetch(
                    `${CONFIG.APPS_SCRIPT_URL}?sheet=${CONFIG.CHECKLIST_SHEET_NAME}&action=fetch`
                )

                if (!response.ok) throw new Error("Failed to fetch checklist data")

                const txt = await response.text()

                try {
                    const parsed = JSON.parse(txt)
                    if (parsed.table && parsed.table.rows) allRows = parsed.table.rows
                    else if (Array.isArray(parsed)) allRows = parsed
                    else if (parsed.values) allRows = parsed.values.map(r => ({ c: r.map(v => ({ v: v })) }))
                } catch (e) {
                    const jsonStart = txt.indexOf("{")
                    const jsonEnd = txt.lastIndexOf("}")
                    if (jsonStart !== -1 && jsonEnd !== -1) {
                        const parsed = JSON.parse(txt.substring(jsonStart, jsonEnd + 1))
                        if (parsed.table && parsed.table.rows) allRows = parsed.table.rows
                    }
                }

                // Get headers from row 0
                headers = allRows[0]?.c?.map(c => c?.v) || []
                setAllChecklistRows(allRows)
                setChecklistHeaders(headers)
            }

            console.log("Checklist Headers:", headers)
            const doerIndex = headers.findIndex(h =>
                h?.toString().toLowerCase().includes('name') ||
                h?.toString().toLowerCase().includes('doer') ||
                h?.toString().toLowerCase().includes('assignee') ||
                h?.toString().toLowerCase().includes('user')
            )

            // Pre-calculate column indices from headers (case-insensitive keyword match)
            // Task ID = Column F (Index 5), Description = Column B (Index 1), Date = Column G (Index 6)
            const taskIdIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l === 'task id' || l === 'id' || l === 'taskid'
                })
                return i !== -1 ? i : 5
            })()
            const descIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l.includes('description') || l.includes('desc') || l.includes('task name') || l.includes('task description') || l === 'task'
                })
                return i !== -1 ? i : 1
            })()
            const dateColIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l.includes('start date') || l.includes('date')
                })
                return i !== -1 ? i : 6
            })()

            const remarksIdx = (() => {
                const i = headers.findIndex(h => {
                    const l = h?.toString().toLowerCase() || ''
                    return l.includes('remark')
                })
                // Fallback to index 13 if not found by name, as seen in handleTransferSubmit
                return i !== -1 ? i : 13
            })()

            console.log("Searching for User:", userName, "| ID col:", taskIdIdx, "| Desc col:", descIdx, "| Date col:", dateColIdx, "| Remarks col:", remarksIdx)

            const userTasks = allRows.slice(1).map((row, idx) => {
                let vals = []
                if (row.c) vals = row.c.map(c => c?.v)
                else if (Array.isArray(row)) vals = row

                const assignee = doerIndex !== -1 ? vals[doerIndex] : vals[5]

                // Match user name + Check if task already has a remark (frontend-only filter)
                // Also ignore tasks that are already "DONE"
                const aName = String(assignee || '').trim().toLowerCase()
                const uName = String(userName || '').trim().toLowerCase()
                // Strict exact match — prevents tasks from other users bleeding through
                const isAssignedToUser = aName !== '' && uName !== '' && aName === uName
                
                const statusIdx = 12 // Standard for this sheet
                const statusValue = String(vals[statusIdx] || '').trim().toUpperCase()
                const isDone = statusValue === 'DONE'

                const remarkValue = remarksIdx !== -1 ? String(vals[remarksIdx] || '').trim() : ''
                const hasNoRemark = remarkValue === ''

                if (isAssignedToUser && !isDone && hasNoRemark) {
                    return {
                        id: vals[taskIdIdx],
                        description: vals[descIdx],
                        date: vals[dateColIdx] || 'N/A',
                        originalRowIndex: idx + 2,
                        allValues: vals
                    }
                }
                return null
            }).filter(t => t !== null)


            setChecklistTasks(userTasks)

        } catch (e) {
            console.error("Error fetching checklist tasks:", e)
        } finally {
            setChecklistLoading(false)
        }
    }

    const handleCheckboxClick = (row) => {
        const headers = leaveData[0]?.headers || []

        const getVal = (name) => {
            const idx = headers.findIndex(h => h?.toLowerCase().includes(name.toLowerCase()))
            return idx !== -1 ? row.values[idx] : ''
        }

        // Extract User Name dynamically from the row headers
        const userName = getVal('name') || getVal('doer') || getVal('assignee') || row.values[3] || "User"

        const taskObj = {
            id: row.id,
            taskId: getVal('task id') || getVal('id') || row.values[0],
            description: getVal('description') || getVal('desc') || getVal('task') || row.values[1],
            date: getVal('date') || getVal('time') || new Date().toLocaleDateString(),
            userName: userName,
            ...row
        }

        const todayStr = new Date().toISOString().split('T')[0]
        setTransferTask(taskObj)
        setTransferForm({
            startDate: todayStr,
            endDate: todayStr,
            isChecked: false
        })
        setSelectedChecklistTaskIds([])
        setShowTransferModal(true)

        if (userName && userName !== "User") {
            fetchUserChecklistTasks(userName)
        }
    }
    // Optimized Transfer Submit: Now uses Optimistic UI for instant closing and local patching
    const handleTransferSubmit = async () => {
        if (selectedChecklistTaskIds.length === 0) {
            alert("Please select at least one task to transfer.")
            return
        }

        try {
            // Pre-calculate data for local patch and JSON payload
            const now = new Date()
            const actStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`
            const fmt = (s) => (s ? `${String(new Date(s).getDate()).padStart(2, '0')}/${String(new Date(s).getMonth() + 1).padStart(2, '0')}/${new Date(s).getFullYear()}` : 'N/A')
            const remarkText = `Leave from ${fmt(transferForm.startDate)} to ${fmt(transferForm.endDate)}`

            const taskMap = new Map(checklistTasks.map(t => [t.id, t]))
            const payload = selectedChecklistTaskIds.map(id => {
                const t = taskMap.get(id)
                return t ? { rowIndex: t.originalRowIndex, taskId: t.id, actualDate: actStr, remarks: remarkText } : null
            }).filter(Boolean)

            // 🚀 STEP 1: OPTIMISTIC PATCH (Make it feel instant)
            if (allChecklistRows) {
                const updatedAllRows = [...allChecklistRows]
                const remIdx = checklistHeaders.findIndex(h => h?.toString().toLowerCase().includes('remark'))
                const actIdx = checklistHeaders.findIndex(h => h?.toString().toLowerCase().includes('actual'))
                
                payload.forEach(pt => {
                    const idx = pt.rowIndex - 1
                    if (updatedAllRows[idx]) {
                        const row = updatedAllRows[idx]
                        const cells = row.c ? [...row.c] : [...row]
                        if (remIdx !== -1) { if (row.c) cells[remIdx] = { v: pt.remarks }; else cells[remIdx] = pt.remarks; }
                        if (actIdx !== -1) { if (row.c) cells[actIdx] = { v: pt.actualDate }; else cells[actIdx] = pt.actualDate; }
                        updatedAllRows[idx] = row.c ? { ...row, c: cells } : cells
                    }
                })
                setAllChecklistRows(updatedAllRows)
            }

            // 🚀 STEP 2: UI FEEDBACK (Instant closing)
            setChecklistTasks([])
            setShowTransferModal(false)
            setSuccessMessage(`Submitting ${payload.length} tasks recorded on leave in background...`)
            setTimeout(() => setSuccessMessage(''), 3000)

            // 🚀 STEP 3: BACKGROUND SUBMIT (No user wait time)
            // Using a slightly larger chunk (60) to avoid excessive boot-up calls for GAS
            const CHUNK_SIZE = 60
            for (let i = 0; i < payload.length; i += CHUNK_SIZE) {
                const chunk = payload.slice(i, i + CHUNK_SIZE)
                fetch(CONFIG.APPS_SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({
                        sheetName: CONFIG.CHECKLIST_SHEET_NAME,
                        action: 'updateTaskData',
                        rowData: JSON.stringify(chunk)
                    })
                }).catch(err => console.error("Background sync failed:", err))
            }

        } catch (e) {
            console.error(e)
            alert("Action failed. Please refresh.")
        }
    }


    // Submit only the Remarks column for a Leave row back to the Unique sheet
    const submitLeaveRemark = async (row) => {
        const headers = leaveData[0]?.headers || []
        const remarksColIndex = headers.findIndex(h => h?.toLowerCase().includes('remark'))
        if (remarksColIndex === -1) {
            alert('Remarks column not found in the Unique sheet headers.')
            return
        }

        const remarkValue = leaveRemarks[row._rowIndex] ?? row.values[remarksColIndex] ?? ''

        try {
            setSavingRemarksId(row._rowIndex)

            // Build a full row copy and replace only the Remarks cell
            const fullRowData = [...row.values]
            while (fullRowData.length <= remarksColIndex) fullRowData.push('')
            fullRowData[remarksColIndex] = remarkValue

            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    sheetName: CONFIG.LEAVE_SHEET_NAME,
                    action: 'update',
                    rowIndex: String(row._rowIndex),
                    rowData: JSON.stringify(fullRowData)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                // Update local leaveData so it reflects the saved value
                setLeaveData(prev => prev.map(r =>
                    r._rowIndex === row._rowIndex
                        ? { ...r, values: fullRowData }
                        : r
                ))
                setSuccessMessage('Remarks saved successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
            } else {
                throw new Error(result.error || 'Failed to save remarks')
            }
        } catch (err) {
            console.error('Error saving remarks:', err)
            alert(`Error: ${err.message}`)
        } finally {
            setSavingRemarksId(null)
        }
    }

    const handleEditClick = (row) => {
        setEditingId(row.id)
        setEditedValues({
            doerName: row.doerName,
            password: row.password,
            role: row.role,
            idEmail: row.idEmail,
            number: row.number
        })
    }

    const handleCancelEdit = () => {
        setEditingId(null)
        setEditedValues({ doerName: '', password: '', role: '', idEmail: '', number: '' })
    }

    const handleSaveEdit = async (row) => {
        if (saving) return
        try {
            setSaving(true)
            setSuccessMessage('')
            const originalRow = row._originalRow || []
            const rowData = [...originalRow]

            rowData[3] = editedValues.doerName
            rowData[4] = editedValues.password
            rowData[5] = editedValues.role
            rowData[6] = editedValues.idEmail
            rowData[7] = editedValues.number

            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    sheetName: CONFIG.SHEET_NAME,
                    action: 'update',
                    rowIndex: String(row._rowIndex),
                    rowData: JSON.stringify(rowData)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                setData(prev => prev.map(item =>
                    item.id === row.id
                        ? {
                            ...item,
                            doerName: editedValues.doerName,
                            password: editedValues.password,
                            role: editedValues.role,
                            idEmail: editedValues.idEmail,
                            number: editedValues.number,
                            _originalRow: rowData
                        }
                        : item
                ))
                setEditingId(null)
                setSuccessMessage('Updated successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
            } else {
                throw new Error(result.error || 'Failed to update')
            }
        } catch (err) {
            console.error("Error updating data:", err)
            alert(`Error: ${err.message}`)
        } finally {
            setSaving(false)
        }
    }

    const handleDeleteRow = async (row) => {
        if (!confirm(`Are you sure you want to delete this entry?\n\nDoer's Name: ${row.doerName}\nPassword: ${row.password}`)) {
            return
        }
        try {
            setDeleting(row.id)
            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    action: 'deleteRow',
                    sheet: CONFIG.SHEET_NAME,
                    rowIndex: String(row._rowIndex)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                await fetchWhatsappData()
                setSuccessMessage('Deleted successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
            } else {
                throw new Error(result.error || 'Failed to delete')
            }
        } catch (err) {
            console.error("Error deleting row:", err)
            alert(`Error: ${err.message}`)
        } finally {
            setDeleting(null)
        }
    }

    const handleAddNewEntry = async () => {
        if (adding) return
        if (!newEntry.doerName.trim()) {
            alert('Doer\'s Name is required')
            return
        }

        try {
            setAdding(true)
            const usersWithData = data.filter(row => row.doerName && row.doerName.trim() !== '')
            const lastUserRowIndex = usersWithData.length > 0
                ? Math.max(...usersWithData.map(row => row._rowIndex))
                : 1
            const nextRowIndex = lastUserRowIndex + 1

            const rowData = ['', '', '', newEntry.doerName, newEntry.password, newEntry.role, newEntry.idEmail, newEntry.number]

            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    sheetName: CONFIG.SHEET_NAME,
                    action: 'update',
                    rowIndex: String(nextRowIndex),
                    rowData: JSON.stringify(rowData)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                setNewEntry({ doerName: '', password: '', role: '', idEmail: '', number: '' })
                setShowAddModal(false)
                setSuccessMessage('New user added successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
                await fetchWhatsappData()
            } else {
                throw new Error(result.error || 'Failed to add user')
            }
        } catch (err) {
            console.error("Error adding user:", err)
            alert(`Error: ${err.message}`)
        } finally {
            setAdding(false)
        }
    }

    const handleAddDepartment = async () => {
        if (addingDept) return
        if (!newDeptEntry.department.trim()) {
            alert('Department is required')
            return
        }

        try {
            setAddingDept(true)
            const deptsWithData = data.filter(row => row.department && row.department.trim() !== '')
            const lastDeptRowIndex = deptsWithData.length > 0
                ? Math.max(...deptsWithData.map(row => row._rowIndex))
                : 1
            const nextRowIndex = lastDeptRowIndex + 1

            const rowData = [newDeptEntry.department, newDeptEntry.givenBy, '', '', '', '', '', '']

            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    sheetName: CONFIG.SHEET_NAME,
                    action: 'update',
                    rowIndex: String(nextRowIndex),
                    rowData: JSON.stringify(rowData)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                setNewDeptEntry({ department: '', givenBy: '' })
                setShowDeptModal(false)
                setSuccessMessage('New department added successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
                await fetchWhatsappData()
            } else {
                throw new Error(result.error || 'Failed to add department')
            }
        } catch (err) {
            console.error("Error adding department:", err)
            alert(`Error: ${err.message}`)
        } finally {
            setAddingDept(false)
        }
    }

    const handleDeptEditClick = (row) => {
        setEditingDeptId(row.id)
        setEditedDeptValues({
            department: row.department,
            givenBy: row.givenBy
        })
    }

    const handleCancelDeptEdit = () => {
        setEditingDeptId(null)
        setEditedDeptValues({ department: '', givenBy: '' })
    }

    const handleSaveDeptEdit = async (row) => {
        if (savingDept) return
        try {
            setSavingDept(true)
            const originalRow = row._originalRow || []
            const rowData = [...originalRow]
            rowData[0] = editedDeptValues.department
            rowData[1] = editedDeptValues.givenBy

            const response = await fetch(CONFIG.APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    sheetName: CONFIG.SHEET_NAME,
                    action: 'update',
                    rowIndex: String(row._rowIndex),
                    rowData: JSON.stringify(rowData)
                })
            })

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
            const result = await response.json()

            if (result.success) {
                // If department name changed, optionally update all users in that department
                // For simplicity, we just refresh local data
                setData(prev => prev.map(item =>
                    item.id === row.id
                        ? {
                            ...item,
                            department: editedDeptValues.department,
                            givenBy: editedDeptValues.givenBy,
                            _originalRow: rowData
                        }
                        : item
                ))
                setEditingDeptId(null)
                setSuccessMessage('Department updated successfully!')
                setTimeout(() => setSuccessMessage(''), 3000)
            } else {
                throw new Error(result.error || 'Failed to update department')
            }
        } catch (err) {
            console.error("Error updating department:", err)
            alert(`Error: ${err.message}`)
        } finally {
            setSavingDept(false)
        }
    }

    // Filter handlers
    const clearUsernameFilter = () => {
        setUsernameFilter('')
        setUsernameDropdownOpen(false)
    }

    const handleUsernameFilterSelect = (name) => {
        setUsernameFilter(name)
        setUsernameDropdownOpen(false)
    }

    const toggleUsernameDropdown = () => {
        setUsernameDropdownOpen(!usernameDropdownOpen)
    }

    const getRoleColor = (role) => {
        const r = role ? role.toLowerCase() : ""
        if (r.includes('admin')) return 'bg-blue-100 text-blue-800'
        if (r.includes('super')) return 'bg-purple-100 text-purple-800'
        return 'bg-gray-100 text-gray-800'
    }

    if (!authorized) return null

    return (
        <AdminLayout>
            <div className="space-y-8 animate-in fade-in duration-500">
                {/* Header and Tabs */}
                <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-4">
                    <div>
                        <h1 className="text-3xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-purple-600 to-pink-600">
                            User Management System
                        </h1>
                        <p className="text-gray-500 mt-1">Manage your organization's users and departments</p>
                    </div>

                    <div className="flex items-center gap-4">
                        <div className="flex border border-purple-200 rounded-md overflow-hidden self-start">
                            <button
                                className={`flex px-4 py-3 text-sm font-medium transition-colors ${activeTab === 'users' ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 hover:bg-purple-50'}`}
                                onClick={() => setActiveTab('users')}
                            >
                                <User size={18} className="mr-2" />
                                Users
                            </button>
                            <button
                                className={`flex px-4 py-3 text-sm font-medium transition-colors ${activeTab === 'departments' ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 hover:bg-purple-50'}`}
                                onClick={() => setActiveTab('departments')}
                            >
                                <Building size={18} className="mr-2" />
                                Departments
                            </button>
                            <button
                                className={`flex px-4 py-3 text-sm font-medium transition-colors ${activeTab === 'leave' ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 hover:bg-purple-50'}`}
                                onClick={() => setActiveTab('leave')}
                            >
                                <LogOut size={18} className="mr-2" />
                                Leave
                            </button>
                            <button
                                className={`flex px-4 py-3 text-sm font-medium transition-colors ${activeTab === 'leave_history' ? 'bg-purple-600 text-white' : 'bg-white text-purple-600 hover:bg-purple-50'}`}
                                onClick={() => setActiveTab('leave_history')}
                            >
                                <History size={18} className="mr-2" />
                                Leave History
                            </button>
                        </div>

                        <button
                            onClick={() => activeTab === 'users' ? setShowAddModal(true) : setShowDeptModal(true)}
                            className="rounded-md bg-gradient-to-r from-blue-600 to-indigo-600 py-2.5 px-6 text-white hover:from-blue-700 hover:to-indigo-800 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 shadow-md transition-all active:scale-95"
                        >
                            <div className="flex items-center">
                                <Plus size={18} className="mr-2" />
                                <span>{activeTab === 'users' ? 'Add User' : 'Add Department'}</span>
                            </div>
                        </button>
                    </div>
                </div>

                {/* Success/Error messages */}
                {successMessage && (
                    <div className="p-4 bg-green-50 border-l-4 border-green-400 text-green-700 rounded-lg shadow-sm animate-in slide-in-from-top duration-300">
                        <div className="flex items-center">
                            <div className="flex-shrink-0">
                                <RefreshCw className="h-5 w-5 text-green-400" />
                            </div>
                            <p className="ml-3 text-sm font-medium">{successMessage}</p>
                        </div>
                    </div>
                )}
                {error && (
                    <div className="p-4 bg-red-50 border-l-4 border-red-400 text-red-700 rounded-lg shadow-sm">
                        <p className="text-sm">{error}</p>
                    </div>
                )}

                {/* Users Tab */}
                {activeTab === 'users' && (
                    <div className="bg-white shadow-xl rounded-2xl overflow-hidden border border-purple-100 transition-all hover:shadow-2xl">
                        <div className="bg-gradient-to-r from-purple-50 to-pink-50 border-b border-purple-100 px-6 py-5 flex justify-between items-center">
                            <div className="flex items-center gap-3">
                                <Users className="text-purple-600" size={20} />
                                <h2 className="text-lg font-bold text-purple-800">User Directory</h2>
                            </div>

                            {/* Username Filter */}
                            <div className="relative flex items-center gap-2">
                                <div className="relative group">
                                    <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-purple-400 group-focus-within:text-purple-600 transition-colors" size={16} />
                                    <input
                                        type="text"
                                        placeholder="Filter by name..."
                                        value={usernameFilter}
                                        onChange={(e) => setUsernameFilter(e.target.value)}
                                        className="w-48 pl-10 pr-8 py-2 border border-purple-200 rounded-full focus:outline-none focus:ring-2 focus:ring-purple-500 text-sm bg-white/80 backdrop-blur-sm transition-all"
                                    />
                                    {usernameFilter && (
                                        <button
                                            onClick={clearUsernameFilter}
                                            className="absolute right-2 top-1/2 transform -translate-y-1/2 text-gray-400 hover:text-purple-600 transition-colors"
                                        >
                                            <X size={14} />
                                        </button>
                                    )}
                                </div>

                                <div className="relative">
                                    <button
                                        onClick={toggleUsernameDropdown}
                                        className="flex items-center gap-1 px-3 py-2 border border-purple-200 rounded-full bg-white text-xs font-semibold text-purple-700 hover:bg-purple-50 hover:border-purple-300 transition-all shadow-sm"
                                    >
                                        Names
                                        <ChevronDown size={14} className={`transition-transform duration-300 ${usernameDropdownOpen ? 'rotate-180' : ''}`} />
                                    </button>

                                    {usernameDropdownOpen && (
                                        <div className="absolute z-50 mt-2 w-56 rounded-xl bg-white shadow-2xl border border-purple-100 max-h-64 overflow-auto top-full right-0 animate-in zoom-in-95 duration-200">
                                            <div className="p-2 space-y-1">
                                                <button
                                                    onClick={clearUsernameFilter}
                                                    className={`block w-full text-left px-4 py-2 text-sm rounded-lg transition-colors ${!usernameFilter ? 'bg-purple-600 text-white' : 'text-gray-700 hover:bg-purple-50'}`}
                                                >
                                                    All Users
                                                </button>
                                                {data.filter(u => u.doerName).map(user => (
                                                    <button
                                                        key={`filter-${user.id}`}
                                                        onClick={() => handleUsernameFilterSelect(user.doerName)}
                                                        className={`block w-full text-left px-4 py-2 text-sm rounded-lg transition-colors ${usernameFilter === user.doerName ? 'bg-purple-100 text-purple-900 border-l-4 border-purple-600' : 'text-gray-700 hover:bg-purple-50'}`}
                                                    >
                                                        {user.doerName}
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    )}
                                </div>
                                <button
                                    onClick={fetchWhatsappData}
                                    className="p-2 text-purple-600 hover:bg-purple-100 rounded-full transition-colors"
                                    title="Refresh Data"
                                >
                                    <RefreshCw size={18} className={loading ? 'animate-spin' : ''} />
                                </button>
                            </div>
                        </div>

                        <div className="overflow-x-auto max-h-[600px] overflow-y-auto scrollbar-thin scrollbar-thumb-purple-200">
                            {loading ? (
                                <div className="flex flex-col justify-center items-center py-24 gap-3">
                                    <Loader2 className="h-12 w-12 text-purple-600 animate-spin" />
                                    <p className="text-purple-600 font-medium animate-pulse">Fetching users...</p>
                                </div>
                            ) : (
                                <table className="min-w-full divide-y divide-purple-100 shadow-sm">
                                    <thead className="bg-purple-50 sticky top-0 z-20 shadow-sm shadow-purple-100/50">
                                        <tr>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Department</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Given By</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Doer's Name</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Password</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Role</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">ID/Email</th>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Number</th>
                                            <th className="px-6 py-4 text-center text-xs font-bold text-purple-700 uppercase tracking-widest">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody className="bg-white divide-y divide-purple-50">
                                        {data
                                            .filter(user =>
                                                (!usernameFilter || user.doerName.toLowerCase().includes(usernameFilter.toLowerCase()))
                                            )
                                            .map((row) => (
                                                <tr key={row.id} className={`hover:bg-purple-50/30 transition-colors ${editingId === row.id ? 'bg-indigo-50/50' : ''}`}>
                                                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-700 font-medium">{row.department || '-'}</td>
                                                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-600">{row.givenBy || '-'}</td>
                                                    <td className="px-6 py-4 whitespace-nowrap">
                                                        {editingId === row.id ? (
                                                            <input
                                                                type="text"
                                                                value={editedValues.doerName}
                                                                onChange={(e) => setEditedValues(prev => ({ ...prev, doerName: e.target.value }))}
                                                                className="w-full px-3 py-1.5 border border-purple-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm shadow-inner"
                                                            />
                                                        ) : (
                                                            <div className="text-sm font-bold text-gray-900">{row.doerName}</div>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 whitespace-nowrap">
                                                        {editingId === row.id ? (
                                                            <input
                                                                type="text"
                                                                value={editedValues.password}
                                                                onChange={(e) => setEditedValues(prev => ({ ...prev, password: e.target.value }))}
                                                                className="w-full px-3 py-1.5 border border-purple-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm font-mono shadow-inner"
                                                            />
                                                        ) : (
                                                            <div className="text-sm text-gray-900 font-mono tracking-wider">{row.password}</div>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 whitespace-nowrap">
                                                        {editingId === row.id ? (
                                                            <select
                                                                value={editedValues.role}
                                                                onChange={(e) => setEditedValues(prev => ({ ...prev, role: e.target.value }))}
                                                                className="w-full px-3 py-1.5 border border-purple-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm"
                                                            >
                                                                <option value="user">user</option>
                                                                <option value="admin">admin</option>
                                                                <option value="super_admin">super_admin</option>
                                                            </select>
                                                        ) : (
                                                            <span className={`px-2.5 py-1 inline-flex text-xs leading-5 font-bold rounded-full shadow-sm capitalize ${getRoleColor(row.role)}`}>
                                                                {row.role}
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 whitespace-nowrap">
                                                        {editingId === row.id ? (
                                                            <input
                                                                type="text"
                                                                value={editedValues.idEmail}
                                                                onChange={(e) => setEditedValues(prev => ({ ...prev, idEmail: e.target.value }))}
                                                                className="w-full px-3 py-1.5 border border-purple-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm shadow-inner"
                                                            />
                                                        ) : (
                                                            <div className="text-sm text-gray-600">{row.idEmail || '-'}</div>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 whitespace-nowrap">
                                                        {editingId === row.id ? (
                                                            <input
                                                                type="text"
                                                                value={editedValues.number}
                                                                onChange={(e) => setEditedValues(prev => ({ ...prev, number: e.target.value }))}
                                                                className="w-full px-3 py-1.5 border border-purple-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm shadow-inner"
                                                            />
                                                        ) : (
                                                            <div className="text-sm text-gray-600">{row.number || '-'}</div>
                                                        )}
                                                    </td>
                                                    <td className="px-6 py-4 whitespace-nowrap text-center">
                                                        {editingId === row.id ? (
                                                            <div className="flex items-center justify-center gap-2">
                                                                <button
                                                                    onClick={() => handleSaveEdit(row)}
                                                                    disabled={saving}
                                                                    className="p-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-all shadow-md active:scale-90"
                                                                    title="Save Changes"
                                                                >
                                                                    {saving ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                                                                </button>
                                                                <button
                                                                    onClick={handleCancelEdit}
                                                                    className="p-2 bg-gray-400 text-white rounded-lg hover:bg-gray-500 transition-all shadow-md active:scale-90"
                                                                    title="Cancel"
                                                                >
                                                                    <X size={16} />
                                                                </button>
                                                            </div>
                                                        ) : (
                                                            <div className="flex items-center justify-center gap-2">
                                                                <button
                                                                    onClick={() => handleEditClick(row)}
                                                                    className="p-2 text-indigo-600 hover:bg-indigo-100 rounded-lg transition-all active:scale-90"
                                                                    title="Edit User"
                                                                >
                                                                    <Edit size={18} />
                                                                </button>
                                                                <button
                                                                    onClick={() => handleDeleteRow(row)}
                                                                    disabled={deleting === row.id}
                                                                    className="p-2 text-red-500 hover:bg-red-50 rounded-lg transition-all active:scale-90"
                                                                    title="Delete User"
                                                                >
                                                                    {deleting === row.id ? <Loader2 size={18} className="animate-spin" /> : <Trash2 size={18} />}
                                                                </button>
                                                            </div>
                                                        )}
                                                    </td>
                                                </tr>
                                            ))}
                                    </tbody>
                                </table>
                            )}
                        </div>
                    </div>
                )}

                {/* Departments Tab */}
                {activeTab === 'departments' && (
                    <div className="bg-white shadow-xl rounded-2xl overflow-hidden border border-purple-100 transition-all hover:shadow-2xl">
                        <div className="bg-gradient-to-r from-purple-50 to-pink-50 border-b border-purple-100 px-6 py-5">
                            <div className="flex justify-between items-center">
                                <div className="flex items-center gap-3">
                                    <Building2 className="text-purple-600" size={20} />
                                    <h2 className="text-lg font-bold text-purple-800">Department Management</h2>
                                </div>

                                <div className="flex border border-purple-200 rounded-full overflow-hidden p-1 bg-white/50 backdrop-blur-sm shadow-inner">
                                    <button
                                        className={`px-6 py-2 text-xs font-bold rounded-full transition-all duration-300 ${activeDeptSubTab === 'departments' ? 'bg-purple-600 text-white shadow-md' : 'text-purple-600 hover:bg-purple-100'}`}
                                        onClick={() => setActiveDeptSubTab('departments')}
                                    >
                                        Names
                                    </button>
                                    <button
                                        className={`px-6 py-2 text-xs font-bold rounded-full transition-all duration-300 ${activeDeptSubTab === 'givenBy' ? 'bg-purple-600 text-white shadow-md' : 'text-purple-600 hover:bg-purple-100'}`}
                                        onClick={() => setActiveDeptSubTab('givenBy')}
                                    >
                                        Assignees
                                    </button>
                                </div>
                            </div>
                        </div>

                        <div className="overflow-x-auto">
                            {loading ? (
                                <div className="flex flex-col justify-center items-center py-24 gap-3">
                                    <Loader2 className="h-12 w-12 text-purple-600 animate-spin" />
                                    <p className="text-purple-600 font-medium">Fetching departments...</p>
                                </div>
                            ) : (
                                <table className="min-w-full divide-y divide-purple-100">
                                    <thead className="bg-purple-50/50">
                                        <tr>
                                            <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest w-20">#</th>
                                            {activeDeptSubTab === 'departments' ? (
                                                <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Department Name</th>
                                            ) : (
                                                <th className="px-6 py-4 text-left text-xs font-bold text-purple-700 uppercase tracking-widest">Given By</th>
                                            )}
                                            <th className="px-6 py-4 text-center text-xs font-bold text-purple-700 uppercase tracking-widest w-32">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody className="bg-white divide-y divide-purple-50">
                                        {data.filter(row => row.department || row.givenBy).map((row, index) => (
                                            <tr key={`dept-${row.id}`} className={`hover:bg-purple-50/30 transition-colors ${editingDeptId === row.id ? 'bg-indigo-50/50' : ''}`}>
                                                <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 font-medium">{index + 1}</td>
                                                <td className="px-6 py-4 whitespace-nowrap">
                                                    {editingDeptId === row.id ? (
                                                        <div className="flex gap-4">
                                                            <div className="flex-1">
                                                                <label className="text-[10px] font-bold text-purple-400 uppercase">Department</label>
                                                                <input
                                                                    type="text"
                                                                    value={editedDeptValues.department}
                                                                    onChange={(e) => setEditedDeptValues(prev => ({ ...prev, department: e.target.value }))}
                                                                    className="w-full px-3 py-2 border border-purple-300 rounded-xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm shadow-inner mt-1"
                                                                />
                                                            </div>
                                                            <div className="flex-1">
                                                                <label className="text-[10px] font-bold text-purple-400 uppercase">Given By</label>
                                                                <input
                                                                    type="text"
                                                                    value={editedDeptValues.givenBy}
                                                                    onChange={(e) => setEditedDeptValues(prev => ({ ...prev, givenBy: e.target.value }))}
                                                                    className="w-full px-3 py-2 border border-purple-300 rounded-xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm shadow-inner mt-1"
                                                                />
                                                            </div>
                                                        </div>
                                                    ) : (
                                                        <div className="text-sm font-bold text-gray-900">
                                                            {activeDeptSubTab === 'departments' ? (row.department || '-') : (row.givenBy || '-')}
                                                        </div>
                                                    )}
                                                </td>
                                                <td className="px-6 py-4 whitespace-nowrap text-center">
                                                    {editingDeptId === row.id ? (
                                                        <div className="flex items-center justify-center gap-2">
                                                            <button
                                                                onClick={() => handleSaveDeptEdit(row)}
                                                                disabled={savingDept}
                                                                className="p-2 bg-green-500 text-white rounded-lg hover:bg-green-600 transition-all shadow-md active:scale-90"
                                                                title="Save Changes"
                                                            >
                                                                {savingDept ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />}
                                                            </button>
                                                            <button
                                                                onClick={handleCancelDeptEdit}
                                                                className="p-2 bg-gray-400 text-white rounded-lg hover:bg-gray-500 transition-all shadow-md active:scale-90"
                                                                title="Cancel"
                                                            >
                                                                <X size={16} />
                                                            </button>
                                                        </div>
                                                    ) : (
                                                        <button
                                                            onClick={() => handleDeptEditClick(row)}
                                                            className="p-2 text-indigo-600 hover:bg-indigo-100 rounded-lg transition-all active:scale-90"
                                                            title="Edit Department"
                                                        >
                                                            <Edit size={18} />
                                                        </button>
                                                    )}
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            )}
                        </div>
                    </div>

                )}

                {/* Leave Tab */}
                {activeTab === 'leave' && (
                    <div className="bg-white shadow-xl rounded-2xl overflow-hidden border border-purple-100 transition-all hover:shadow-2xl">
                        <div className="bg-gradient-to-r from-purple-50 to-pink-50 border-b border-purple-100 px-6 py-5 flex justify-between items-center">
                            <div className="flex items-center gap-3">
                                <LogOut className="text-purple-600" size={20} />
                                <h2 className="text-lg font-bold text-purple-800">Leave Management</h2>
                            </div>
                            <div className="flex items-center gap-4">
                                <div className="relative group">
                                    <User size={14} className="absolute left-3 top-1/2 transform -translate-y-1/2 text-purple-400 pointer-events-none" />
                                    <select
                                        value={leaveNameFilter}
                                        onChange={(e) => setLeaveNameFilter(e.target.value)}
                                        className="pl-9 pr-8 py-1.5 border border-purple-100 rounded-xl focus:outline-none focus:ring-2 focus:ring-purple-500 text-[11px] bg-white/50 transition-all shadow-sm appearance-none cursor-pointer font-bold text-gray-700 min-w-[160px]"
                                    >
                                        <option value="">All Staff Members</option>
                                        {(() => {
                                            const headers = leaveData[0]?.headers || []
                                            const nameIdx = headers.findIndex(h => {
                                                const l = h?.toLowerCase() || ''
                                                return l.includes('name') || l.includes('doer') || l.includes('user')
                                            })
                                            const idx = nameIdx !== -1 ? nameIdx : 3
                                            const names = [...new Set(leaveData.slice(1).map(row => row.values[idx]?.toString().trim()).filter(Boolean))].sort()
                                            return names.map(name => (
                                                <option key={name} value={name}>{name}</option>
                                            ))
                                        })() || null}
                                    </select>
                                    <div className="absolute right-2 top-1/2 transform -translate-y-1/2 pointer-events-none text-purple-400">
                                        <ChevronDown size={14} />
                                    </div>
                                </div>
                                <button
                                    onClick={fetchLeaveData}
                                    className="p-2 text-purple-600 hover:bg-purple-100 rounded-full transition-colors"
                                    title="Refresh Data"
                                >
                                    <RefreshCw size={18} className={leaveLoading ? 'animate-spin' : ''} />
                                </button>
                            </div>
                        </div>

                        <div className="overflow-x-auto overflow-y-auto" style={{ maxHeight: '60vh' }}>
                            {leaveLoading ? (
                                <div className="flex flex-col justify-center items-center py-24 gap-3">
                                    <Loader2 className="h-12 w-12 text-purple-600 animate-spin" />
                                    <p className="text-purple-600 font-medium animate-pulse">Fetching leave data...</p>
                                </div>
                            ) : leaveError ? (
                                <div className="p-8 text-center text-red-500">
                                    <p>Error loading data: {leaveError}</p>
                                    <button onClick={fetchLeaveData} className="mt-4 px-4 py-2 bg-red-100 rounded-lg hover:bg-red-200 text-red-700 font-semibold">Retry</button>
                                </div>
                            ) : leaveData.length === 0 ? (
                                <div className="p-12 text-center text-gray-500">
                                    <p>No records found in Unique sheet.</p>
                                </div>
                            ) : (() => {
                                return (
                                    <table className="min-w-full divide-y divide-gray-200">
                                        <thead className="bg-gray-50 sticky top-0 z-10">
                                            <tr>
                                                <th className="px-3 py-3 w-10 text-center">
                                                    <span className="sr-only">Leave</span>
                                                </th>
                                                <th className="px-4 py-3 text-left text-[10px] font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                                                    Department
                                                </th>
                                                <th className="px-4 py-3 text-left text-[10px] font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                                                    Given By
                                                </th>
                                                <th className="px-4 py-3 text-left text-[10px] font-bold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                                                    Name
                                                </th>
                                            </tr>
                                        </thead>
                                        <tbody className="bg-white divide-y divide-gray-100">
                                            {uniqueLeaveRows
                                                .filter(row => {
                                                    if (!leaveNameFilter) return true
                                                    const headers = leaveData[0]?.headers || []
                                                    const nameIdx = headers.findIndex(h => {
                                                        const l = h?.toLowerCase() || ''
                                                        return l.includes('name') || l.includes('doer') || l.includes('user')
                                                    })
                                                    const idx = nameIdx !== -1 ? nameIdx : 3
                                                    return row.values[idx]?.toString().trim() === leaveNameFilter
                                                })
                                                .map((row, rowIndex) => {
                                                    const headers = leaveData[0]?.headers || []
                                                    const deptIdx = (() => {
                                                        const i = headers.findIndex(h => h?.toLowerCase().includes('dept') || h?.toLowerCase().includes('department'))
                                                        return i !== -1 ? i : 0
                                                    })()

                                                    const givenByIdx = (() => {
                                                        const i = headers.findIndex(h => h?.toLowerCase().includes('given') || h?.toLowerCase().includes('assignee'))
                                                        return i !== -1 ? i : 1
                                                    })()

                                                    const nameIdx = (() => {
                                                        const i = headers.findIndex(h => {
                                                            const l = h?.toLowerCase() || ''
                                                            return l.includes('name') || l.includes('doer') || l.includes('user')
                                                        })
                                                        return i !== -1 ? i : 3
                                                    })()

                                                    const dept = row.values[deptIdx] || '—'
                                                    const givenBy = row.values[givenByIdx] || '—'
                                                    const name = row.values[nameIdx] || '—'

                                                    return (
                                                        <tr
                                                            key={row.id || rowIndex}
                                                            className="hover:bg-purple-50/30 transition-colors group"
                                                        >
                                                            {/* Checkbox — opens transfer/leave modal */}
                                                            <td className="px-3 py-3 text-center whitespace-nowrap">
                                                                <input
                                                                    type="checkbox"
                                                                    className="rounded border-gray-300 text-purple-600 focus:ring-purple-500 cursor-pointer"
                                                                    onClick={(e) => {
                                                                        e.stopPropagation()
                                                                        handleCheckboxClick(row)
                                                                    }}
                                                                />
                                                            </td>
                                                            <td className="px-4 py-3 text-xs text-gray-700 font-medium whitespace-nowrap">{dept}</td>
                                                            <td className="px-4 py-3 text-xs text-gray-700 whitespace-nowrap">{givenBy}</td>
                                                            <td className="px-4 py-3 text-xs text-gray-900 font-semibold whitespace-nowrap">{name}</td>
                                                        </tr>
                                                    )
                                                })}
                                        </tbody>
                                    </table>
                                )
                            })()}
                        </div>
                    </div>
                )}

                {/* Leave History Tab */}
                {activeTab === 'leave_history' && (() => {
                    // Date parser for filtering
                    const parseDate = (dStr) => {
                        if (!dStr) return null
                        const s = String(dStr).trim()
                        // Google Sheets Date(y,m,d)
                        const gsMatch = s.match(/^Date\((\d+),(\d+),(\d+)\)$/)
                        if (gsMatch) return new Date(Number(gsMatch[1]), Number(gsMatch[2]), Number(gsMatch[3]))
                        // DD/MM/YYYY
                        const dmyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
                        if (dmyMatch) return new Date(`${dmyMatch[3]}-${dmyMatch[2].padStart(2, '0')}-${dmyMatch[1].padStart(2, '0')}`)
                        const d = new Date(s)
                        return isNaN(d.getTime()) ? null : d
                    }

                    const start = leaveHistoryStartDate ? new Date(leaveHistoryStartDate) : null
                    const end = leaveHistoryEndDate ? new Date(leaveHistoryEndDate) : null
                    if (end) end.setHours(23, 59, 59, 999)

                    const filteredRaw = leaveHistoryData.filter(row => {
                        // Filter by Name (exact match from dropdown)
                        if (leaveHistoryNameFilter && row.name !== leaveHistoryNameFilter) return false
                        
                        // Filter by Date
                        if (start || end) {
                            const taskDate = parseDate(row.date)
                            if (!taskDate) return false
                            if (start && taskDate < start) return false
                            if (end && taskDate > end) return false
                        }
                        return true
                    })

                    // Deduplicate by Date (Day only) and Name
                    const filtered = []
                    const seen = new Set()
                    for (const row of filteredRaw) {
                        const dObj = parseDate(row.date)
                        if (dObj) dObj.setHours(0, 0, 0, 0)
                        const d = dObj ? dObj.getTime() : String(row.date).trim().toLowerCase()
                        const n = String(row.name || '').trim().replace(/\s+/g, ' ').toLowerCase()
                        const key = `${d}|${n}`
                        if (!seen.has(key)) {
                            seen.add(key)
                            filtered.push(row)
                        }
                    }

                    const uniqueNames = [...new Set(leaveHistoryData.map(row => row.name).filter(Boolean))].sort()

                    return (
                        <div className="bg-white shadow-xl rounded-2xl overflow-hidden border border-purple-100 transition-all hover:shadow-2xl">
                            <div className="bg-gradient-to-r from-purple-50 to-pink-50 border-b border-purple-100 px-6 py-5 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                                <div className="flex items-center gap-3">
                                    <div className="p-2 bg-purple-600 rounded-lg shadow-lg">
                                        <History className="text-white" size={20} />
                                    </div>
                                    <div>
                                        <div className="flex items-center gap-3">
                                            <h2 className="text-lg font-bold text-purple-800">Leave History Analytics</h2>
                                            {!leaveHistoryLoading && !leaveHistoryError && (
                                                <span className="px-2.5 py-0.5 bg-purple-100 text-purple-700 text-[10px] font-bold rounded-full border border-purple-200 uppercase tracking-tighter">
                                                    {filtered.length} {filtered.length === 1 ? 'Record' : 'Records'}
                                                </span>
                                            )}
                                        </div>
                                        <p className="text-[10px] text-purple-400 font-medium tracking-tight">Track and analyze historical leave data</p>
                                    </div>
                                </div>

                                <div className="flex flex-wrap items-center gap-3">
                                    {/* Name Filter Dropdown */}
                                    <div className="relative group">
                                        <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-purple-400 group-focus-within:text-purple-600 transition-colors pointer-events-none" size={16} />
                                        <select
                                            value={leaveHistoryNameFilter}
                                            onChange={(e) => setLeaveHistoryNameFilter(e.target.value)}
                                            className="w-44 pl-9 pr-8 py-2 border border-purple-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-purple-500 text-xs bg-white transition-all shadow-sm appearance-none cursor-pointer font-medium text-gray-700"
                                        >
                                            <option value="">All Doers</option>
                                            {uniqueNames.map(name => (
                                                <option key={name} value={name}>{name}</option>
                                            ))}
                                        </select>
                                        <div className="absolute right-3 top-1/2 transform -translate-y-1/2 pointer-events-none text-purple-400 group-focus-within:text-purple-600 transition-colors">
                                            <ChevronDown size={14} />
                                        </div>
                                    </div>

                                    {/* Date Range */}
                                    <div className="flex items-center gap-2">
                                        <div className="relative">
                                            <Calendar className="absolute left-3 top-1/2 transform -translate-y-1/2 text-purple-400" size={14} />
                                            <input
                                                type="date"
                                                value={leaveHistoryStartDate}
                                                onChange={(e) => setLeaveHistoryStartDate(e.target.value)}
                                                className="pl-9 pr-3 py-2 border border-purple-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-purple-500 text-xs bg-white shadow-sm"
                                            />
                                        </div>
                                        <span className="text-purple-400 text-sm font-medium">to</span>
                                        <div className="relative">
                                            <Calendar className="absolute left-3 top-1/2 transform -translate-y-1/2 text-purple-400" size={14} />
                                            <input
                                                type="date"
                                                value={leaveHistoryEndDate}
                                                onChange={(e) => setLeaveHistoryEndDate(e.target.value)}
                                                className="pl-9 pr-3 py-2 border border-purple-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-purple-500 text-xs bg-white shadow-sm"
                                            />
                                        </div>
                                    </div>

                                    <button
                                        onClick={fetchLeaveHistoryData}
                                        className="p-2 text-purple-600 hover:bg-purple-600 hover:text-white rounded-xl transition-all bg-white border border-purple-100 shadow-sm active:scale-95"
                                        title="Refresh History"
                                    >
                                        <RefreshCw size={18} className={leaveHistoryLoading ? 'animate-spin' : ''} />
                                    </button>
                                </div>
                            </div>

                            <div className="overflow-x-auto overflow-y-auto" style={{ maxHeight: '60vh' }}>
                                {leaveHistoryLoading ? (
                                    <div className="flex flex-col justify-center items-center py-24 gap-3">
                                        <div className="relative">
                                            <Loader2 className="h-12 w-12 text-purple-600 animate-spin" />
                                            <History className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 text-purple-300" size={16} />
                                        </div>
                                        <p className="text-purple-600 font-bold animate-pulse text-xs uppercase tracking-widest">Analyzing history...</p>
                                    </div>
                                ) : leaveHistoryError ? (
                                    <div className="p-16 text-center">
                                        <div className="inline-flex p-4 rounded-full bg-red-50 text-red-500 mb-4 border border-red-100">
                                            <X size={24} />
                                        </div>
                                        <h3 className="text-red-800 font-bold mb-2">Fetch Failed</h3>
                                        <p className="text-gray-500 mb-6 text-sm max-w-xs mx-auto">Error loading history: {leaveHistoryError}</p>
                                        <button onClick={fetchLeaveHistoryData} className="px-8 py-2.5 bg-gradient-to-r from-purple-600 to-pink-600 text-white rounded-xl font-bold shadow-lg shadow-purple-200 hover:opacity-90 transition-all active:scale-95 text-xs">Retry Now</button>
                                    </div>
                                ) : filtered.length === 0 ? (
                                    <div className="p-24 text-center">
                                        <div className="inline-flex p-4 rounded-full bg-purple-50 text-purple-300 mb-4">
                                            <History size={32} />
                                        </div>
                                        <p className="text-gray-400 text-sm italic">No leave records match your current filters.</p>
                                        {(leaveHistoryNameFilter || leaveHistoryStartDate || leaveHistoryEndDate) && (
                                            <button 
                                                onClick={() => {
                                                    setLeaveHistoryNameFilter('')
                                                    setLeaveHistoryStartDate('')
                                                    setLeaveHistoryEndDate('')
                                                }}
                                                className="mt-4 text-xs font-bold text-purple-600 hover:underline"
                                            >
                                                Clear all filters
                                            </button>
                                        )}
                                    </div>
                                ) : (() => {
                                    // Robust Date Formatter for Table Display
                                    const formatTableDate = (val) => {
                                        const d = parseDate(val)
                                        if (!d) return String(val || '—')
                                        const day = String(d.getDate()).padStart(2, '0')
                                        const month = String(d.getMonth() + 1).padStart(2, '0')
                                        const year = d.getFullYear()
                                        return `${day}/${month}/${year}`
                                    }

                                    return (
                                        <table className="min-w-full divide-y divide-purple-100">
                                            <thead className="bg-purple-50/50 sticky top-0 z-20 backdrop-blur-md">
                                                <tr>
                                                    <th className="px-4 py-2.5 text-left text-[10px] font-bold text-purple-700 uppercase tracking-widest border-b border-purple-100/50">Date</th>
                                                    <th className="px-4 py-2.5 text-left text-[10px] font-bold text-purple-700 uppercase tracking-widest border-b border-purple-100/50">Doer Name</th>
                                                    <th className="px-4 py-2.5 text-left text-[10px] font-bold text-purple-700 uppercase tracking-widest border-b border-purple-100/50">Leave Details (Remarks)</th>
                                                </tr>
                                            </thead>
                                            <tbody className="bg-white divide-y divide-purple-50">
                                                {filtered.map((row, index) => (
                                                    <tr key={`${row.date}-${row.name}-${index}`} className="hover:bg-purple-50/40 transition-all duration-200 group">
                                                        <td className="px-4 py-1.5 whitespace-nowrap text-[11px] text-gray-500 font-mono tracking-tight">{formatTableDate(row.date)}</td>
                                                        <td className="px-4 py-1.5 whitespace-nowrap text-xs font-bold text-gray-800 group-hover:text-purple-700 transition-colors">{row.name || '—'}</td>
                                                        <td className="px-4 py-1.5 text-xs text-gray-700 italic border-l-4 border-transparent group-hover:border-purple-400 group-hover:bg-purple-50/50 transition-all pl-4">
                                                            {row.remarks}
                                                        </td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )
                                })()}
                            </div>
                        </div>
                    )
                })()}
            </div>

            {/* Modals with Updated Styling */}
            {/* User Modal */}
            {
                showAddModal && (
                    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
                        <div className="fixed inset-0 bg-purple-900/40 backdrop-blur-sm animate-in fade-in duration-300" onClick={() => setShowAddModal(false)}></div>
                        <div className="relative bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in-95 duration-300">
                            <div className="bg-gradient-to-r from-purple-600 to-pink-600 px-8 py-6 flex justify-between items-center">
                                <div>
                                    <h2 className="text-xl font-bold text-white">Create New User</h2>
                                    <p className="text-purple-100 text-xs mt-1">Add a new doer to the system</p>
                                </div>
                                <button onClick={() => setShowAddModal(false)} className="p-2 bg-white/10 hover:bg-white/20 rounded-full transition-colors">
                                    <X className="text-white" size={20} />
                                </button>
                            </div>

                            <div className="p-8 space-y-5">
                                <div className="grid grid-cols-2 gap-5">
                                    <div className="space-y-1.5">
                                        <label className="text-xs font-bold text-purple-600 uppercase ml-1">Doer's Name *</label>
                                        <input
                                            type="text"
                                            value={newEntry.doerName}
                                            onChange={(e) => setNewEntry(prev => ({ ...prev, doerName: e.target.value }))}
                                            placeholder="Enter name"
                                            className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 transition-all"
                                        />
                                    </div>
                                    <div className="space-y-1.5">
                                        <label className="text-xs font-bold text-purple-600 uppercase ml-1">Password</label>
                                        <input
                                            type="text"
                                            value={newEntry.password}
                                            onChange={(e) => setNewEntry(prev => ({ ...prev, password: e.target.value }))}
                                            placeholder="Enter password"
                                            className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm font-mono bg-purple-50/30 transition-all"
                                        />
                                    </div>
                                </div>

                                <div className="grid grid-cols-2 gap-5">
                                    <div className="space-y-1.5">
                                        <label className="text-xs font-bold text-purple-600 uppercase ml-1">Role</label>
                                        <select
                                            value={newEntry.role}
                                            onChange={(e) => setNewEntry(prev => ({ ...prev, role: e.target.value }))}
                                            className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 appearance-none transition-all"
                                        >
                                            <option value="">Select Role</option>
                                            <option value="user">User</option>
                                            <option value="admin">Admin</option>
                                            <option value="super_admin">Super Admin</option>
                                        </select>
                                    </div>
                                    <div className="space-y-1.5">
                                        <label className="text-xs font-bold text-purple-600 uppercase ml-1">ID / Email</label>
                                        <input
                                            type="text"
                                            value={newEntry.idEmail}
                                            onChange={(e) => setNewEntry(prev => ({ ...prev, idEmail: e.target.value }))}
                                            placeholder="ID or Email"
                                            className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 transition-all"
                                        />
                                    </div>
                                </div>

                                <div className="space-y-1.5">
                                    <label className="text-xs font-bold text-purple-600 uppercase ml-1">Phone Number</label>
                                    <input
                                        type="text"
                                        value={newEntry.number}
                                        onChange={(e) => setNewEntry(prev => ({ ...prev, number: e.target.value }))}
                                        placeholder="Enter number"
                                        className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 transition-all"
                                    />
                                </div>

                                <div className="pt-4 flex gap-4">
                                    <button
                                        onClick={() => setShowAddModal(false)}
                                        className="flex-1 py-3 px-6 rounded-2xl border border-purple-100 font-bold text-purple-600 hover:bg-purple-50 transition-all active:scale-95"
                                    >
                                        Cancel
                                    </button>
                                    <button
                                        onClick={handleAddNewEntry}
                                        disabled={adding}
                                        className="flex-[2] py-3 px-6 rounded-2xl bg-gradient-to-r from-purple-600 to-pink-600 text-white font-bold shadow-lg hover:shadow-purple-200 transition-all active:scale-95 flex justify-center items-center"
                                    >
                                        {adding ? <Loader2 size={18} className="animate-spin" /> : 'Create User'}
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            }

            {/* Department Modal */}
            {
                showDeptModal && (
                    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
                        <div className="fixed inset-0 bg-purple-900/40 backdrop-blur-sm animate-in fade-in duration-300" onClick={() => setShowDeptModal(false)}></div>
                        <div className="relative bg-white rounded-3xl shadow-2xl w-full max-w-md overflow-hidden animate-in zoom-in-95 duration-300">
                            <div className="bg-gradient-to-r from-purple-600 to-pink-600 px-8 py-6 flex justify-between items-center">
                                <div>
                                    <h2 className="text-xl font-bold text-white">New Department</h2>
                                    <p className="text-purple-100 text-xs mt-1">Add a functional group</p>
                                </div>
                                <button onClick={() => setShowDeptModal(false)} className="p-2 bg-white/10 hover:bg-white/20 rounded-full transition-colors">
                                    <X className="text-white" size={20} />
                                </button>
                            </div>

                            <div className="p-8 space-y-5">
                                <div className="space-y-1.5">
                                    <label className="text-xs font-bold text-purple-600 uppercase ml-1">Department Name *</label>
                                    <input
                                        type="text"
                                        value={newDeptEntry.department}
                                        onChange={(e) => setNewDeptEntry(prev => ({ ...prev, department: e.target.value }))}
                                        placeholder="e.g. Sales, HR"
                                        className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 transition-all"
                                    />
                                </div>

                                <div className="space-y-1.5">
                                    <label className="text-xs font-bold text-purple-600 uppercase ml-1">Given By (Assignee)</label>
                                    <input
                                        type="text"
                                        value={newDeptEntry.givenBy}
                                        onChange={(e) => setNewDeptEntry(prev => ({ ...prev, givenBy: e.target.value }))}
                                        placeholder="Enter name"
                                        className="w-full px-4 py-3 border border-purple-100 rounded-2xl focus:ring-2 focus:ring-purple-500 focus:outline-none text-sm bg-purple-50/30 transition-all"
                                    />
                                </div>

                                <div className="pt-4 flex gap-4">
                                    <button
                                        onClick={() => setShowDeptModal(false)}
                                        className="flex-1 py-3 px-6 rounded-2xl border border-purple-100 font-bold text-purple-600 hover:bg-purple-50 transition-all active:scale-95"
                                    >
                                        Cancel
                                    </button>
                                    <button
                                        onClick={handleAddDepartment}
                                        disabled={addingDept}
                                        className="flex-[2] py-3 px-6 rounded-2xl bg-gradient-to-r from-purple-600 to-pink-600 text-white font-bold shadow-lg hover:shadow-purple-200 transition-all active:scale-95 flex justify-center items-center"
                                    >
                                        {addingDept ? <Loader2 size={18} className="animate-spin" /> : 'Create Department'}
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            }

            {/* Transfer/Leave Modal */}
            {showTransferModal && transferTask && (() => {
                // Formatting helpers for the modal
                const parseDateLocal = (s) => {
                    if (!s || s === 'N/A') return null
                    // Handle Google Sheets Date format or DD/MM/YYYY
                    const gsMatch = s.toString().match(/^Date\((\d+),(\d+),(\d+)\)$/)
                    if (gsMatch) return new Date(Number(gsMatch[1]), Number(gsMatch[2]), Number(gsMatch[3]))
                    const dmyMatch = s.toString().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
                    if (dmyMatch) return new Date(`${dmyMatch[3]}-${dmyMatch[2].padStart(2, '0')}-${dmyMatch[1].padStart(2, '0')}`)
                    const d = new Date(s)
                    return isNaN(d.getTime()) ? null : d
                }

                // 🚀 OPTIMIZATION: Calculation once per modal render
                const filterStart = parseDateLocal(transferForm.startDate)
                const filterEnd = parseDateLocal(transferForm.endDate)
                if (filterEnd) filterEnd.setHours(23, 59, 59, 999)

                const modalTasks = checklistTasks.filter(task => {
                    if (task.userName !== transferForm.userName) return false
                    if (!filterStart && !filterEnd) return true
                    const taskDate = parseDateLocal(task.startDate || task.date)
                    if (!taskDate) return false
                    if (filterStart && filterEnd) return taskDate >= filterStart && taskDate <= filterEnd
                    if (filterStart) return taskDate >= filterStart
                    if (filterEnd) return taskDate <= filterEnd
                    return true
                })

                const selectedSet = new Set(selectedChecklistTaskIds)
                const allSelected = modalTasks.length > 0 && modalTasks.every(t => selectedSet.has(t.id))

                return (
                    <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
                        <div className="fixed inset-0 bg-gray-900/50 backdrop-blur-sm" onClick={() => setShowTransferModal(false)}></div>
                        <div className="relative bg-white rounded-xl shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in-95 duration-300 flex flex-col" style={{ maxHeight: '90vh' }}>
                            {/* Header */}
                            <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-white">
                                <h2 className="text-lg font-bold text-gray-800">
                                    Transfer Tasks for <span className="text-purple-600">{transferForm.userName || "User"}</span>
                                </h2>
                                <button onClick={() => setShowTransferModal(false)} className="text-gray-400 hover:text-gray-600">
                                    <X size={20} />
                                </button>
                            </div>

                            {/* Body */}
                            <div className="p-6 space-y-6 overflow-y-auto flex-1">
                                <div className="grid grid-cols-2 gap-4">
                                    <div>
                                        <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">From Date</label>
                                        <input
                                            type="date"
                                            value={transferForm.startDate}
                                            onChange={(e) => setTransferForm(prev => ({ ...prev, startDate: e.target.value }))}
                                            className="w-full px-3 py-2 border rounded-lg text-sm focus:ring-2 focus:ring-purple-500"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">To Date</label>
                                        <input
                                            type="date"
                                            value={transferForm.endDate}
                                            onChange={(e) => setTransferForm(prev => ({ ...prev, endDate: e.target.value }))}
                                            className="w-full px-3 py-2 border rounded-lg text-sm focus:ring-2 focus:ring-purple-500"
                                        />
                                    </div>
                                </div>

                                <div className="space-y-3">
                                    <div className="flex items-center justify-between">
                                        <div className="flex items-center space-x-2">
                                            <input
                                                type="checkbox"
                                                id="selAllMod"
                                                className="w-4 h-4 rounded border-gray-300 text-purple-600"
                                                checked={allSelected}
                                                onChange={(e) => {
                                                    if (e.target.checked) {
                                                        const newIds = Array.from(new Set([...selectedChecklistTaskIds, ...modalTasks.map(t => t.id)]))
                                                        setSelectedChecklistTaskIds(newIds)
                                                    } else {
                                                        const visibleIds = new Set(modalTasks.map(t => t.id))
                                                        setSelectedChecklistTaskIds(prev => prev.filter(id => !visibleIds.has(id)))
                                                    }
                                                }}
                                            />
                                            <label htmlFor="selAllMod" className="text-xs font-bold text-gray-700 uppercase cursor-pointer">
                                                Select Available ({modalTasks.length})
                                            </label>
                                        </div>
                                    </div>

                                    <div className="border border-gray-100 rounded-lg overflow-hidden" style={{ maxHeight: '340px' }}>
                                        <table className="min-w-full divide-y divide-gray-100">
                                            <thead className="bg-gray-50 sticky top-0 z-10">
                                                <tr>
                                                    <th className="w-10"></th>
                                                    <th className="px-4 py-2 text-left text-[10px] font-bold text-gray-400 uppercase">ID</th>
                                                    <th className="px-4 py-2 text-left text-[10px] font-bold text-gray-400 uppercase">Task</th>
                                                    <th className="px-4 py-2 text-left text-[10px] font-bold text-gray-400 uppercase">Date</th>
                                                </tr>
                                            </thead>
                                            <tbody className="bg-white divide-y divide-gray-50">
                                                {checklistLoading ? (
                                                    <tr><td colSpan={4} className="py-8 text-center"><Loader2 className="animate-spin h-5 w-5 mx-auto text-purple-500" /></td></tr>
                                                ) : modalTasks.length === 0 ? (
                                                    <tr><td colSpan={4} className="py-8 text-center text-xs text-gray-400">No tasks match criteria.</td></tr>
                                                ) : (
                                                    modalTasks.map((task) => (
                                                        <tr key={task.id} className={`hover:bg-purple-50/30 transition-colors pointer-events-auto cursor-pointer ${selectedSet.has(task.id) ? 'bg-purple-50' : ''}`}
                                                            onClick={() => {
                                                                setSelectedChecklistTaskIds(prev =>
                                                                    selectedSet.has(task.id) ? prev.filter(id => id !== task.id) : [...prev, task.id]
                                                                )
                                                            }}
                                                        >
                                                            <td className="px-4 py-3 text-center">
                                                                <input type="checkbox" checked={selectedSet.has(task.id)} readOnly className="rounded border-gray-300 text-purple-600" />
                                                            </td>
                                                            <td className="px-4 py-3 text-xs font-mono text-gray-400">{task.id}</td>
                                                            <td className="px-4 py-3 text-xs text-gray-600 max-w-xs truncate">{task.description}</td>
                                                            <td className="px-4 py-3 text-xs text-gray-400">{task.startDate || task.date}</td>
                                                        </tr>
                                                    ))
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>

                            {/* Footer */}
                            <div className="px-6 py-4 bg-gray-50 border-t flex justify-between items-center">
                                <span className="text-xs font-bold text-gray-400">{selectedChecklistTaskIds.length} Selected</span>
                                <div className="flex gap-2">
                                    <button onClick={() => setShowTransferModal(false)} className="px-4 py-2 text-xs font-bold text-gray-500 hover:text-gray-700">Cancel</button>
                                    <button
                                        onClick={handleTransferSubmit}
                                        disabled={selectedChecklistTaskIds.length === 0}
                                        className="px-6 py-2 bg-purple-600 text-white rounded-lg text-xs font-bold hover:bg-purple-700 disabled:bg-gray-200 transition-colors"
                                    >
                                        Apply Leave
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            })()}
        </AdminLayout >
    )
}

export default Settings