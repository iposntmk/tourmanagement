const { createApp, ref, reactive, computed, watch } = Vue;

const App = {
    setup() {
        // STATE
        const activeTab = ref('tours');
        const filterMonth = ref('all');
        const editingItem = ref(null); // For inline editing
        const editModal = reactive({ target: null, data: {} });

        const pasteTexts = reactive({
            companies: '', nationalities: '', provinces: '', expenseTypes: '', locations: '', detailedExpenses: ''
        });
        
        const masterData = reactive({
            companies: [
                { id: 1, name: 'Vietravel' },
                { id: 2, name: 'Saigontourist' },
                { id: 3, name: 'Hanoi Tourist' }
            ],
            nationalities: [
                { id: 1, name: 'Việt Nam' },
                { id: 2, name: 'Hoa Kỳ' },
                { id: 3, name: 'Hàn Quốc' }
            ],
            provinces: [
                { id: 1, name: 'Quảng Bình', order: 1 },
                { id: 2, name: 'Quảng Trị', order: 2 },
                { id: 3, name: 'Huế', order: 3 },
                { id: 4, name: 'Đà Nẵng', order: 4 },
                { id: 5, name: 'Hội An', order: 5 },
                { id: 6, name: 'Mỹ Sơn', order: 6 },
            ],
            locations: [
                { id: 101, provinceId: 3, name: 'Đại Nội', price: 200000 },
                { id: 102, provinceId: 3, name: 'Lăng Khải Định', price: 150000 },
                { id: 103, provinceId: 4, name: 'Ngũ Hành Sơn', price: 40000 },
                { id: 104, provinceId: 4, name: 'Bà Nà Hills', price: 850000 },
                { id: 105, provinceId: 1, name: 'Động Phong Nha', price: 150000 },
                { id: 106, provinceId: 5, name: 'Phố cổ Hội An', price: 120000 },
            ],
            expenseTypes: [
                { id: 1, name: 'Ăn uống' },
                { id: 2, name: 'Di chuyển' },
                { id: 3, name: 'Lưu trú' },
                { id: 4, name: 'Dịch vụ khác'}
            ],
            detailedExpenses: [
                { id: 1, expenseTypeId: 1, name: 'Ăn trưa', amount: 150000 },
                { id: 2, expenseTypeId: 1, name: 'Ăn tối', amount: 200000 },
                { id: 3, expenseTypeId: 2, name: 'Taxi', amount: 100000 },
                { id: 4, expenseTypeId: 4, name: 'Nước uống', amount: 10000 },
                { id: 5, expenseTypeId: 4, name: 'Đạp xe cùng khách', amount: 50000 },
                { id: 6, expenseTypeId: 4, name: 'Đi ăn tối riêng', amount: 100000 },
                { id: 7, expenseTypeId: 4, name: 'Xem show', amount: 250000 },
            ]
        });

        const tours = reactive([
             { id: 'TOUR001', originalId: 'TOUR001', companyId: 1, guide: 'Cao Hữu Tú', nationalityId: 1, startDate: '2025-07-10', endDate: '2025-07-12', paxAdult: 2, paxChild: 1, notes: 'Khách VIP', tickets: [{ locationId: 101 }, { locationId: 102 }], operatingCosts: [{ id: Date.now(), detailedExpenseId: 1, quantity: 2, price: 150000}], fees: { workDaysQB: 0, workDaysOther: 2, overnightStays: 1, airportPickup: true }, incidentals: [{ id: 1, name: 'Cà phê', amount: 50000, notes: '' }], payment: { advance: 3000000, collect: 0 } }
        ]);

        const currentTour = ref(null);
        const newItems = reactive({
            company: '',
            nationality: '',
            province: '',
            location: { name: '', provinceId: '', price: 0 },
            expenseType: '',
            detailedExpense: { name: '', expenseTypeId: '', amount: 0 }
        });

        const ticketForm = reactive({
            provinceId: '',
            locationId: ''
        });

        const serviceForm = reactive({
            detailedExpenseId: '',
            quantity: 1
        });
        
        const toasts = ref([]);

        // COMPUTED PROPERTIES
        const tourDuration = computed(() => {
            if (!currentTour.value || !currentTour.value.startDate || !currentTour.value.endDate) return 0;
            const start = new Date(currentTour.value.startDate);
            const end = new Date(currentTour.value.endDate);
            if (isNaN(start) || isNaN(end) || end < start) return 0;
            const diffTime = Math.abs(end - start);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
            return diffDays;
        });
        
        const totalWorkDays = computed(() => {
            if (!currentTour.value) return 0;
            return (currentTour.value.fees.workDaysQB || 0) + (currentTour.value.fees.workDaysOther || 0);
        });

        const workDaysExceeded = computed(() => {
            if (!currentTour.value) return false;
            return totalWorkDays.value > tourDuration.value;
        });

        const totalOperatingCost = computed(() => {
            if (!currentTour.value?.operatingCosts) return 0;
            return currentTour.value.operatingCosts.reduce((sum, item) => {
                const effectiveQuantity = item.quantity;
                return sum + (item.price * effectiveQuantity);
            }, 0);
        });

        const filteredTours = computed(() => {
            if (filterMonth.value === 'all') {
                return tours;
            }
            return tours.filter(t => new Date(t.startDate).getMonth() + 1 == filterMonth.value);
        });

        const monthlyTotal = computed(() => {
            if (filterMonth.value === 'all') return 0;
            return filteredTours.value.reduce((sum, tour) => sum + grandTotalForTour(tour), 0);
        });

        const totalPax = computed(() => (currentTour.value?.paxAdult || 0) + (currentTour.value?.paxChild || 0));

        const filteredLocations = computed(() => {
            if (!ticketForm.provinceId) return [];
            return masterData.locations.filter(l => l.provinceId == ticketForm.provinceId);
        });
        
        const sortedProvinces = computed(() => [...masterData.provinces].sort((a, b) => a.order - b.order));
        
        const sortedLocations = computed(() => {
            const provinceOrder = masterData.provinces.reduce((acc, p) => ({...acc, [p.id]: p.order }), {});
            return [...masterData.locations].sort((a, b) => provinceOrder[a.provinceId] - provinceOrder[b.provinceId]);
        });
        
        const totalTicketCost = computed(() => {
            if (!currentTour.value) return 0;
            return currentTour.value.tickets.reduce((sum, ticket) => {
                const location = masterData.locations.find(l => l.id === ticket.locationId);
                return sum + (location ? location.price * totalPax.value : 0);
            }, 0);
        });

        const totalFeeCost = computed(() => {
            if (!currentTour.value) return 0;
            const fees = currentTour.value.fees;
            return (fees.workDaysQB * 800000) + (fees.workDaysOther * 600000) + (fees.overnightStays * 150000);
        });

        const totalIncidentalCost = computed(() => {
            if (!currentTour.value) return 0;
            return currentTour.value.incidentals.reduce((sum, item) => sum + Number(item.amount || 0), 0);
        });

        const grandTotal = computed(() => {
            if (!currentTour.value) return 0;
            return totalTicketCost.value + totalOperatingCost.value + totalFeeCost.value + totalIncidentalCost.value;
        });
        
        const balanceAfterAdvance = computed(() => {
             if (!currentTour.value) return 0;
             return grandTotal.value - currentTour.value.payment.advance;
        });

        const balanceAfterCollection = computed(() => {
             if (!currentTour.value) return 0;
             return balanceAfterAdvance.value - currentTour.value.payment.collect;
        });
        
        const finalSettlement = computed(() => {
            if (!currentTour.value) return 0;
            return currentTour.value.payment.advance - (grandTotal.value - currentTour.value.payment.collect);
        });

        watch(totalPax, (newVal) => {
            serviceForm.quantity = newVal > 0 ? newVal : 1;
        });

        // METHODS
        const showToast = (message, type = 'success') => {
            const toastId = Date.now();
            const newToast = {
                id: toastId,
                message,
                type: type === 'success' ? 'alert-success' : 'alert-error',
                icon: type === 'success' ? 'fas fa-check-circle' : 'fas fa-exclamation-circle',
            };
            toasts.value.push(newToast);
            setTimeout(() => {
                const index = toasts.value.findIndex(t => t.id === toastId);
                if (index > -1) {
                    toasts.value.splice(index, 1);
                }
            }, 4000);
        };

        const getNewId = (list) => (list.length > 0 ? Math.max(...list.map(i => i.id)) : 0) + 1;

        const getEmptyTour = () => ({
            id: '',
            originalId: null,
            companyId: masterData.companies[0]?.id || '',
            guide: 'Cao Hữu Tú',
            nationalityId: masterData.nationalities[0]?.id || '',
            startDate: new Date().toISOString().split('T')[0],
            endDate: new Date().toISOString().split('T')[0],
            paxAdult: 2,
            paxChild: 0,
            notes: '',
            tickets: [],
            operatingCosts: [],
            fees: { workDaysQB: 0, workDaysOther: 0, overnightStays: 0 },
            incidentals: [],
            payment: { advance: 0, collect: 0 }
        });

        const showTourForm = (tourData) => {
            if (tourData.id) {
                currentTour.value = JSON.parse(JSON.stringify(tourData));
                if (!currentTour.value.operatingCosts) {
                    currentTour.value.operatingCosts = [];
                }
                currentTour.value.originalId = tourData.id;
            } else {
                currentTour.value = getEmptyTour();
            }
        };

        const saveTour = () => {
            if (!currentTour.value.id) {
                showToast('Mã tour không được để trống', 'error');
                return;
            }
            if (!currentTour.value.companyId || !currentTour.value.nationalityId) {
                showToast('Vui lòng chọn Công ty và Quốc tịch', 'error');
                return;
            }
            
            if (currentTour.value.originalId && currentTour.value.originalId !== currentTour.value.id) {
                 const existing = tours.find(t => t.id === currentTour.value.id);
                 if (existing) {
                     showToast(`Mã tour '${currentTour.value.id}' đã tồn tại.`, 'error');
                     return;
                 }
            }

            if (currentTour.value.originalId) {
                const index = tours.findIndex(t => t.id === currentTour.value.originalId);
                if (index !== -1) {
                    tours[index] = JSON.parse(JSON.stringify(currentTour.value));
                    tours[index].originalId = currentTour.value.id;
                }
            } else {
                 const existing = tours.find(t => t.id === currentTour.value.id);
                 if (existing) {
                     showToast(`Mã tour '${currentTour.value.id}' đã tồn tại.`, 'error');
                     return;
                 }
                currentTour.value.originalId = currentTour.value.id;
                tours.unshift(JSON.parse(JSON.stringify(currentTour.value)));
            }
            showToast('Lưu tour thành công!');
            currentTour.value = null;
        };

        const deleteTour = (tourId) => {
            const index = tours.findIndex(t => t.id === tourId);
            if (index !== -1) {
                tours.splice(index, 1);
                showToast('Đã xóa tour ' + tourId);
            }
        };
        
        const copyTour = (tourToCopy) => {
            const newTour = JSON.parse(JSON.stringify(tourToCopy));
            let baseId = tourToCopy.id.split('-COPY')[0];
            let newId = `${baseId}-COPY`;
            let counter = 1;
            while (tours.some(t => t.id === newId)) {
                counter++;
                newId = `${baseId}-COPY-${counter}`;
            }
            newTour.id = newId;
            newTour.originalId = null;
            tours.unshift(newTour);
            showToast(`Đã sao chép tour ${tourToCopy.id} thành ${newId}`);
        };

        const addTicket = () => {
            if (!ticketForm.locationId) return;
            if (currentTour.value.tickets.some(t => t.locationId === ticketForm.locationId)) {
                showToast('Địa điểm này đã được thêm', 'error');
                return;
            }
            currentTour.value.tickets.push({ locationId: ticketForm.locationId });
            
            const provinceOrder = masterData.provinces.reduce((acc, p) => ({...acc, [p.id]: p.order }), {});
            const locationProvinceMap = masterData.locations.reduce((acc, l) => ({...acc, [l.id]: l.provinceId }), {});

            currentTour.value.tickets.sort((a,b) => {
                return provinceOrder[locationProvinceMap[a.locationId]] - provinceOrder[locationProvinceMap[b.locationId]];
            });

            ticketForm.locationId = '';
        };
        
        const addOperatingCost = () => {
            if (!serviceForm.detailedExpenseId) {
                showToast('Vui lòng chọn một chi phí.', 'error');
                return;
            }
            const expense = masterData.detailedExpenses.find(s => s.id === serviceForm.detailedExpenseId);
            if (!expense) return;

            currentTour.value.operatingCosts.push({
                id: Date.now(),
                detailedExpenseId: expense.id,
                quantity: serviceForm.quantity,
                price: expense.amount // Snapshot the price
            });

            serviceForm.detailedExpenseId = '';
            serviceForm.quantity = 1;
        };

        const copyOperatingCost = (itemToCopy, index) => {
            const newItem = JSON.parse(JSON.stringify(itemToCopy));
            newItem.id = Date.now();
            currentTour.value.operatingCosts.splice(index + 1, 0, newItem);
        };

        const copyIncidental = (itemToCopy, index) => {
            const newItem = JSON.parse(JSON.stringify(itemToCopy));
            newItem.id = Date.now();
            currentTour.value.incidentals.splice(index + 1, 0, newItem);
        };

        const isLastInProvince = (index) => {
            const tickets = currentTour.value.tickets;
            if (index === tickets.length - 1) return false;
            const getLocationProvince = (locationId) => masterData.locations.find(l => l.id === locationId)?.provinceId;
            const currentProvince = getLocationProvince(tickets[index].locationId);
            const nextProvince = getLocationProvince(tickets[index + 1].locationId);
            return currentProvince !== nextProvince;
        };

        const formatCurrency = (value) => {
            if (value === null || value === undefined || isNaN(Number(value))) return '0';
            return new Intl.NumberFormat('vi-VN').format(value);
        };

        const formatCurrencyWithK = (value) => {
            if (value === null || value === undefined || isNaN(Number(value))) return '0';
            const numValue = Number(value);
            if (numValue >= 1000) {
                const result = numValue / 1000;
                return (result % 1 === 0 ? result : result.toFixed(1)) + 'k';
            }
            return numValue.toString();
        };

        const formatInputCurrency = (event, item = null, key = null) => {
            let value = event.target.value;
            if (value.toLowerCase().endsWith('k')) {
                value = parseFloat(value.slice(0, -1)) * 1000;
            } else {
                value = value.replace(/\D/g, '');
            }
            const numValue = Number(value) || 0;
            if (item && key) item[key] = numValue;
            event.target.value = formatCurrency(numValue);
        };
        
        const addMasterDataItem = (listName, data, showFeedback = true) => {
            if (!data.name) {
                if (showFeedback) showToast('Tên không được để trống', 'error');
                return false;
            }
            const newItem = { id: getNewId(masterData[listName]), ...data };
            if (listName === 'provinces') newItem['order'] = newItem.id;
            masterData[listName].push(newItem);
            
            // Reset input
            const singleName = listName.slice(0, -1);
            if(newItems[singleName] !== undefined && typeof newItems[singleName] === 'string') newItems[singleName] = '';
            if(newItems.location?.name !== undefined && listName === 'locations') newItems.location = { name: '', provinceId: '', price: 0 };
            if(newItems.detailedExpense?.name !== undefined && listName === 'detailedExpenses') newItems.detailedExpense = { name: '', expenseTypeId: '', amount: 0 };

            if (showFeedback) showToast('Thêm thành công!', 'success');
            return true;
        };

        const copyMasterDataItem = (itemToCopy, listName) => {
            const list = masterData[listName];
            if (!list) return;

            const newItem = JSON.parse(JSON.stringify(itemToCopy));
            newItem.id = getNewId(list);
            newItem.name = `${itemToCopy.name}-COPY`;

            const originalIndex = list.findIndex(item => item.id === itemToCopy.id);
            if (originalIndex !== -1) {
                list.splice(originalIndex + 1, 0, newItem);
            } else {
                list.push(newItem);
            }
            showToast(`Đã sao chép: ${itemToCopy.name}`);
        };
        
        const deleteMasterDataItem = (listName, id) => {
            const index = masterData[listName].findIndex(i => i.id === id);
            if (index !== -1) {
                masterData[listName].splice(index, 1);
                showToast('Xóa thành công', 'success');
            }
        };

        const startEditing = (item, listName) => {
            editingItem.value = { id: item.id, text: item.name, list: listName };
        };

        const saveEdit = () => {
            if (!editingItem.value) return;
            const { id, text, list } = editingItem.value;
            const itemToUpdate = masterData[list].find(i => i.id === id);
            if (itemToUpdate) {
                itemToUpdate.name = text;
            }
            editingItem.value = null;
        };

        const cancelEdit = () => {
            editingItem.value = null;
        };

        const openEditModal = (item, target) => {
            editModal.target = target;
            editModal.data = JSON.parse(JSON.stringify(item)); // Deep copy
            window.edit_modal.showModal();
        };

        const saveEditModal = () => {
            const { target, data } = editModal;
            const index = masterData[target].findIndex(item => item.id === data.id);
            if (index !== -1) {
                masterData[target][index] = data;
                showToast('Cập nhật thành công!');
            }
            window.edit_modal.close();
        };

        const importFromPastedText = (targetList) => {
            const text = pasteTexts[targetList];
            if (!text.trim()) {
                showToast('Không có nội dung để import.', 'error');
                return;
            }
            const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
            let successCount = 0;
            let errorCount = 0;

            try {
                switch (targetList) {
                    case 'companies':
                    case 'nationalities':
                    case 'provinces':
                    case 'expenseTypes':
                        lines.forEach(line => {
                            const name = line.trim();
                            if (name && !masterData[targetList].some(item => item.name.toLowerCase() === name.toLowerCase())) {
                                addMasterDataItem(targetList, {name}, false);
                                successCount++;
                            }
                        });
                        break;
                    
                    case 'locations':
                        lines.forEach(line => {
                            const parts = line.split(',').map(p => p.trim());
                            if (parts.length === 3) {
                                const [name, provinceId, price] = parts;
                                if (name && !isNaN(Number(provinceId)) && !isNaN(Number(price))) {
                                    if (!masterData.locations.some(l => l.name.toLowerCase() === name.toLowerCase())) {
                                        if (addMasterDataItem('locations', { name, provinceId: Number(provinceId), price: Number(price) }, false)) {
                                            successCount++;
                                        } else { errorCount++; }
                                    }
                                } else { errorCount++; }
                            } else { errorCount++; }
                        });
                        break;
                    
                    case 'detailedExpenses':
                        lines.forEach(line => {
                            const parts = line.split(',').map(p => p.trim());
                            if (parts.length === 3) {
                                const [name, expenseTypeId, amount] = parts;
                                if (name && !isNaN(Number(expenseTypeId)) && !isNaN(Number(amount))) {
                                    if (!masterData.detailedExpenses.some(de => de.name.toLowerCase() === name.toLowerCase())) {
                                        if (addMasterDataItem('detailedExpenses', { name, expenseTypeId: Number(expenseTypeId), amount: Number(amount) }, false)) {
                                            successCount++;
                                        } else { errorCount++; }
                                    }
                                } else { errorCount++; }
                            } else { errorCount++; }
                        });
                        break;
                }
                let message = `Import hoàn tất.`;
                if (successCount > 0) message += ` Đã thêm ${successCount} mục mới.`;
                if (errorCount > 0) message += ` ${errorCount} dòng bị lỗi.`;
                showToast(message, errorCount > 0 && successCount === 0 ? 'error' : 'success');
                pasteTexts[targetList] = ''; // Clear textarea after import
            } catch (err) {
                console.error("Import error:", err);
                showToast('Đã xảy ra lỗi trong quá trình import.', 'error');
            }
        };
        
        const getCompanyName = (id) => masterData.companies.find(c => c.id === id)?.name || 'N/A';
        const getProvinceName = (id) => masterData.provinces.find(p => p.id === id)?.name || 'N/A';
        const getLocationName = (id) => masterData.locations.find(l => l.id === id)?.name || 'N/A';
        const getLocationPrice = (id) => masterData.locations.find(l => l.id === id)?.price || 0;
        const getExpenseTypeName = (id) => masterData.expenseTypes.find(et => et.id === id)?.name || 'N/A';
        const getDetailedExpenseName = (id) => masterData.detailedExpenses.find(de => de.id === id)?.name || 'N/A';
        const formatDate = (dateString) => {
            if (!dateString) return 'N/A';
            const date = new Date(dateString);
            return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`;
        };

        const grandTotalForTour = (tour) => {
            const tourTotalPax = tour.paxAdult + tour.paxChild;
            
            const tickets = (tour.tickets || []).reduce((sum, ticket) => {
                const price = getLocationPrice(ticket.locationId);
                return sum + (price * tourTotalPax);
            }, 0);
            
            const operating = (tour.operatingCosts || []).reduce((sum, item) => sum + (item.price * item.quantity), 0);

            const fees = (tour.fees.workDaysQB * 800000) + (tour.fees.workDaysOther * 600000) + (tour.fees.overnightStays * 150000);
            const incidentals = (tour.incidentals || []).reduce((sum, item) => sum + Number(item.amount || 0), 0);

            return tickets + operating + fees + incidentals;
        };

        const sendTourEmail = (tour) => {
            exportToExcel([tour], tour.id);
            showToast('File Excel đã được tải. Vui lòng đính kèm vào email.', 'success');
            const subject = encodeURIComponent(tour.id);
            const body = encodeURIComponent('việt á tour');
            window.location.href = `mailto:huutu289@gmail.com?subject=${subject}&body=${body}`;
        };

        const exportToExcel = (toursToExport, fileName) => {
            const wb = XLSX.utils.book_new();
            const ws_data = [];

            const borderAll = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
            const headerStyle = { font: { bold: true, color: { rgb: "FFFFFFFF" } }, fill: { fgColor: { rgb: "FF4F4F4F" } }, alignment: { horizontal: 'center', vertical: 'center' }, border: borderAll };
            const subHeaderStyleLeft = { font: { bold: true }, alignment: { horizontal: 'left' }, border: borderAll };
            const subHeaderStyleCenter = { font: { bold: true }, alignment: { horizontal: 'center' }, border: borderAll };
            const textCellStyle = { alignment: { horizontal: 'left' }, border: borderAll };
            const numberCellStyle = { alignment: { horizontal: 'right' }, border: borderAll };
            const currencyCellStyle = { alignment: { horizontal: 'right' }, numFmt: '#,##0', border: borderAll };
            const titleStyle = { font: { bold: true, sz: 16, color: {rgb: "FF1E40AF"} } };
            const summaryTitleStyle = { font: { bold: true, sz: 14, color: {rgb: "FFC00000"} } };
            const separatorStyle = { fill: { fgColor: { rgb: "FFD3D3D3" } } };

            let monthlyTotals = { cost: 0, advance: 0, collect: 0, settlement: 0 };

            toursToExport.forEach((tour, tourIndex) => {
                const tourTotalPax = tour.paxAdult + tour.paxChild;

                ws_data.push([{v: `CHI TIẾT TOUR: ${tour.id}`, s: titleStyle }]);
                ws_data.push([{v: 'Mã Tour', s: subHeaderStyleLeft}, {v: tour.id, s: textCellStyle}]);
                ws_data.push([{v: 'Công ty', s: subHeaderStyleLeft}, {v: getCompanyName(tour.companyId), s: textCellStyle}]);
                ws_data.push([{v: 'Hướng dẫn viên', s: subHeaderStyleLeft}, {v: tour.guide, s: textCellStyle}]);
                ws_data.push([{v: 'Ngày', s: subHeaderStyleLeft}, {v: `${formatDate(tour.startDate)} - ${formatDate(tour.endDate)}`, s: textCellStyle}]);
                ws_data.push([{v: 'Số khách', s: subHeaderStyleLeft}, {v: `${tourTotalPax} (${tour.paxAdult} NL, ${tour.paxChild} TE)`, s: textCellStyle}]);
                ws_data.push([]);

                const addSection = (title, headers, dataRows) => {
                    ws_data.push([{v: title, s: headerStyle}, null, null, null]);
                    ws_data.push(headers.map(h => ({v: h, s: subHeaderStyleCenter})));
                    dataRows.forEach(row => {
                        ws_data.push([
                            {v: row[0], s: textCellStyle},
                            {v: row[1], t: 'n', s: numberCellStyle},
                            {v: row[2], t: 'n', s: currencyCellStyle},
                            {v: row[3], t: 'n', s: currencyCellStyle}
                        ]);
                    });
                    ws_data.push([]);
                }

                const operatingCostRows = [];
                (tour.tickets || []).forEach(ticket => {
                    const price = getLocationPrice(ticket.locationId);
                    operatingCostRows.push([`Vé: ${getLocationName(ticket.locationId)}`, tourTotalPax, price, price * tourTotalPax]);
                });
                (tour.operatingCosts || []).forEach(item => {
                    operatingCostRows.push([getDetailedExpenseName(item.detailedExpenseId), item.quantity, item.price, item.price * item.quantity]);
                });
                (tour.incidentals || []).forEach(item => {
                    operatingCostRows.push([item.name, 1, item.amount, item.amount]);
                });
                addSection('Tổng Chi Phí Vận Hành (Vé, Dịch vụ, Phát sinh)', ['Hạng mục', 'Số lượng', 'Đơn giá', 'Thành tiền'], operatingCostRows);

                const feeRows = [];
                if(tour.fees.workDaysQB > 0) feeRows.push(['Công tác ngày (QB)', tour.fees.workDaysQB, 800000, tour.fees.workDaysQB * 800000]);
                if(tour.fees.workDaysOther > 0) feeRows.push(['Công tác ngày (Khác)', tour.fees.workDaysOther, 600000, tour.fees.workDaysOther * 600000]);
                if(tour.fees.overnightStays > 0) feeRows.push(['Lưu đêm/xe', tour.fees.overnightStays, 150000, tour.fees.overnightStays * 150000]);
                addSection('Công tác phí', ['Hạng mục', 'Số lượng', 'Đơn giá', 'Thành tiền'], feeRows);

                const tourGrandTotal = grandTotalForTour(tour);
                const tourFinalSettlement = tour.payment.advance - (tourGrandTotal - tour.payment.collect);
                ws_data.push([{v: 'Tổng kết Tour', s: headerStyle}, null, null, null]);
                ws_data.push([{v: 'Tổng chi phí', s: subHeaderStyleLeft}, '', '', {v: tourGrandTotal, t:'n', s: currencyCellStyle}]);
                ws_data.push([{v: 'Tạm ứng', s: subHeaderStyleLeft}, '', '', {v: tour.payment.advance, t:'n', s: currencyCellStyle}]);
                ws_data.push([{v: 'Thu hộ', s: subHeaderStyleLeft}, '', '', {v: tour.payment.collect, t:'n', s: currencyCellStyle}]);
                ws_data.push([{v: tourFinalSettlement >= 0 ? 'HDV hoàn lại CTY' : 'CTY trả thêm HDV', s: subHeaderStyleLeft}, '', '', {v: Math.abs(tourFinalSettlement), t:'n', s: currencyCellStyle}]);
                
                monthlyTotals.cost += tourGrandTotal;
                monthlyTotals.advance += tour.payment.advance;
                monthlyTotals.collect += tour.payment.collect;
                monthlyTotals.settlement += tourFinalSettlement;

                if (tourIndex < toursToExport.length - 1) {
                    ws_data.push([]); 
                    ws_data.push([{v: '', s: separatorStyle}, {v: '', s: separatorStyle}, {v: '', s: separatorStyle}, {v: '', s: separatorStyle}]);
                    ws_data.push([]);
                }
            });

            if (toursToExport.length > 1) {
                ws_data.push([]);
                ws_data.push([]);
                ws_data.push([{v: `TỔNG KẾT THÁNG ${filterMonth.value}`, s: summaryTitleStyle}]);
                ws_data.push([{v: 'Tổng chi phí tháng', s: subHeaderStyleLeft}, '', '', {v: monthlyTotals.cost, t: 'n', s: currencyCellStyle}]);
                ws_data.push([{v: 'Tổng tạm ứng tháng', s: subHeaderStyleLeft}, '', '', {v: monthlyTotals.advance, t: 'n', s: currencyCellStyle}]);
                ws_data.push([{v: 'Tổng thu hộ tháng', s: subHeaderStyleLeft}, '', '', {v: monthlyTotals.collect, t: 'n', s: currencyCellStyle}]);
                ws_data.push([{v: 'Tổng quyết toán tháng', s: subHeaderStyleLeft}, '', '', {v: monthlyTotals.settlement, t: 'n', s: currencyCellStyle}]);
            }

            const ws = XLSX.utils.aoa_to_sheet(ws_data, {cellStyles: true});
            
            const merges = [];
            for (let i = 0; i < ws_data.length; i++) {
                if (ws_data[i][0]?.s === headerStyle || ws_data[i][0]?.s === summaryTitleStyle || ws_data[i][0]?.s === titleStyle) {
                    merges.push({ s: { r: i, c: 0 }, e: { r: i, c: 3 } });
                }
            }
            ws['!merges'] = merges;

            ws['!cols'] = [{wch: 35}, {wch: 15}, {wch: 15}, {wch: 20}];
            XLSX.utils.book_append_sheet(wb, ws, "ChiTietTour");
            XLSX.writeFile(wb, `${fileName}.xlsx`);
        };

        return {
            activeTab, filterMonth, masterData, tours, currentTour, newItems, ticketForm, toasts, 
            editingItem, editModal, pasteTexts, serviceForm,
            filteredTours, totalPax, filteredLocations, sortedProvinces, sortedLocations,
            totalTicketCost, totalFeeCost, totalIncidentalCost, grandTotal,
            tourDuration, totalOperatingCost,
            totalWorkDays, workDaysExceeded, balanceAfterAdvance, balanceAfterCollection, finalSettlement,
            monthlyTotal,
            showToast, showTourForm, saveTour, deleteTour, copyTour, addTicket, addOperatingCost, copyOperatingCost, copyIncidental, isLastInProvince,
            formatCurrency, formatCurrencyWithK, formatInputCurrency, addMasterDataItem, copyMasterDataItem,
            deleteMasterDataItem, getCompanyName, getProvinceName, getLocationName, getLocationPrice,
            getExpenseTypeName, getDetailedExpenseName, formatDate, 
            startEditing, saveEdit, cancelEdit,
            openEditModal, saveEditModal,
            importFromPastedText, exportToExcel, grandTotalForTour, sendTourEmail
        };
    }
};

createApp(App).mount('#app');
</script>

</body>
</html>
