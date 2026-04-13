// แนะนำให้บันทึกไฟล์นี้ในชื่อ admin-email.js ไว้ในโฟลเดอร์เดียวกันกับ index.html

export function setupAdminEmail(app, db, appId, predefinedUsers, firebaseHelpers) {
    // รับตัวแปร Firebase ที่จำเป็นมาจากไฟล์หลัก เพื่อไม่ให้เกิด Error (Scope)
    const { doc, setDoc, writeBatch } = firebaseHelpers;

    // นำฟังก์ชันทั้งหมดไปผูกกับ window.app ของไฟล์หลัก
    Object.assign(app, {
        // ==============================================
        // Email System Logic & Auto-Pilot
        // ==============================================

        startAutoPilot: function() {
            if (this.autoPilotInterval) clearInterval(this.autoPilotInterval);
            this.autoPilotInterval = setInterval(() => this.checkAndRunAutoPilot(), 60000);
            console.log("AutoPilot System Started");
        },

        stopAutoPilot: function() {
            if (this.autoPilotInterval) {
                clearInterval(this.autoPilotInterval);
                this.autoPilotInterval = null;
            }
            console.log("AutoPilot System Stopped");
        },

        // ฟังก์ชันใหม่: เช็คว่าข้อความนี้ "ค้างส่ง" สำหรับคนๆ นี้หรือไม่ (แยกรายบุคคล)
        isMsgEligibleForUser: function(m, uid) {
            // เช็คก่อนว่าข้อความนี้ส่งถึงคนนี้ไหม
            const isToUser = m.to && Array.isArray(m.to) && (m.to.includes(uid) || m.to.length >= predefinedUsers.length);
            if (!isToUser) return false;
            
            // ถ้าระบุว่าให้รีเซ็ตสถานะเฉพาะคนนี้
            if (m.resetForUsers && m.resetForUsers.includes(uid)) return true;
            
            // คืนค่าสถานะปกติ (ถ้ายังไม่เคยส่งเลย)
            return !m.emailSent;
        },

        checkAndRunAutoPilot: async function() {
            if (!this.profile || !this.profile.isAdmin || !this.emailConfig.autoPilot) return;

            // ป้องกันการรันซ้อนทับกัน ถ้ารอบเก่ายังส่งไม่เสร็จ (แก้ปัญหาส่งเบิ้ล 2-3 รอบ)
            if (this.isProcessingAutoPilot) return;
            this.isProcessingAutoPilot = true;

            try {
                const now = new Date();
                const currentDay = now.getDay().toString(); // 0=Sun, 1=Mon...
                const currentHours = String(now.getHours()).padStart(2, '0');
                const currentMinutes = String(now.getMinutes()).padStart(2, '0');
                const currentTimeStr = `${currentHours}:${currentMinutes}`;

                let msgsToProcess = this.allMessages.filter(m => {
                    // ข้ามข้อความที่เคยส่งแล้วและไม่ได้ถูกสั่งให้รีเซ็ต
                    if (m.emailSent && (!m.resetForUsers || m.resetForUsers.length === 0)) return false;
                    return true;
                });

                if (msgsToProcess.length === 0) return;

                let eligibleMsgs = [];
                let shouldSend = false;

                if (this.emailConfig.sendType === 'delay') {
                    const delayMs = (this.emailConfig.delayMinutes || 0) * 60 * 1000;
                    eligibleMsgs = msgsToProcess.filter(m => (Date.now() - m.timestamp) >= delayMs);
                    if (eligibleMsgs.length > 0) shouldSend = true;
                } else if (this.emailConfig.sendType === 'schedule') {
                    if (this.emailConfig.scheduleDays.includes(currentDay) && this.emailConfig.scheduleTime === currentTimeStr) {
                        // ป้องกันการส่งซ้ำในนาทีเดียวกัน (เผื่อ setInterval รัน 2 รอบใน 1 นาที)
                        const todayStr = now.toDateString() + '-' + currentTimeStr;
                        if (this.lastScheduleSentTime === todayStr) return;
                        this.lastScheduleSentTime = todayStr;

                        eligibleMsgs = msgsToProcess;
                        if (eligibleMsgs.length > 0) shouldSend = true;
                    }
                }

                if (!shouldSend || eligibleMsgs.length === 0) return;

                const userGroups = {};
                // จัดกลุ่มข้อความที่จะส่ง โดยอิงความถูกต้องแบบรายบุคคล
                predefinedUsers.forEach(u => {
                    const userMsgs = eligibleMsgs.filter(m => this.isMsgEligibleForUser(m, u.studentId));
                    if (userMsgs.length > 0) {
                        userGroups[u.studentId] = userMsgs;
                    }
                });

                for (const uid in userGroups) {
                    if ((this.emailConfig.excludedAutoPilotUsers || []).includes(uid)) continue;

                    const user = predefinedUsers.find(u => u.studentId === uid);
                    if (user) {
                        const userMsgs = userGroups[uid]; 
                        const htmlBody = this.buildEmailTemplate(user, userMsgs);
                        const email = `${uid}@kmitl.ac.th`;

                        try {
                            const response = await fetch(this.EMAIL_API_URL, {
                                method: "POST",
                                headers: { "Content-Type": "text/plain;charset=utf-8" },
                                body: JSON.stringify({
                                    email: email,
                                    subject: `คุณมีข้อความใหม่จากระบบ Song Siang [อัตโนมัติ]`,
                                    htmlBody: htmlBody
                                })
                            });
                            
                            const resJson = await response.json();
                            
                            // ถ้าส่งผ่านแล้ว ให้เคลียร์สถานะของคนๆ นี้ออก
                            if (resJson.status === 'success') {
                                const batch = writeBatch(db);
                                userMsgs.forEach(m => {
                                    const ref = doc(db, 'artifacts', appId, 'public', 'data', 'song_siang_messages', m.id);
                                    let newReset = (m.resetForUsers || []).filter(id => id !== uid);
                                    batch.update(ref, { emailSent: true, resetForUsers: newReset });
                                });
                                await batch.commit();
                            }
                        } catch(e) { console.error("AutoPilot send failed", e); }
                        await new Promise(r => setTimeout(r, 1500));
                    }
                }
            } finally {
                this.isProcessingAutoPilot = false;
            }
        },

        setEmailSendType: function(type) {
            this.emailConfig.sendType = type;
            document.getElementById('emailDelaySettings').className = type === 'delay' ? 'block mb-4 p-4 bg-black/10 rounded-xl border border-white/10' : 'hidden';
            document.getElementById('emailScheduleSettings').className = type === 'schedule' ? 'block mb-4 p-4 bg-black/10 rounded-xl border border-white/10' : 'hidden';
            
            document.querySelectorAll('[onclick^="app.setEmailSendType"]').forEach(el => {
                el.className = `flex-1 py-2 rounded-xl text-xs font-bold transition-colors border ${el.getAttribute('onclick').includes(type) ? 'bg-secondary text-primary border-secondary' : 'bg-transparent text-white border-white/40'}`;
            });
        },

        toggleScheduleDay: function(dayStr) {
            const idx = this.emailConfig.scheduleDays.indexOf(dayStr);
            if (idx > -1) {
                this.emailConfig.scheduleDays.splice(idx, 1);
            } else {
                this.emailConfig.scheduleDays.push(dayStr);
            }
            const btn = document.getElementById(`btnDay${dayStr}`);
            if (btn) {
                btn.className = `px-3 py-1 rounded-lg text-xs font-bold transition-colors border ${idx === -1 ? 'bg-secondary text-primary border-secondary' : 'bg-white/10 text-white/60 border-white/30'}`;
            }
        },
        
        toggleAutoPilotCollapse: function() {
            this.isAutoPilotExpanded = !this.isAutoPilotExpanded;
            const content = document.getElementById('autoPilotSettingsContent');
            const icon = document.getElementById('autoPilotCollapseIcon');
            if (content && icon) {
                if (this.isAutoPilotExpanded) {
                    content.classList.remove('hidden');
                    icon.classList.replace('fa-chevron-down', 'fa-chevron-up');
                } else {
                    content.classList.add('hidden');
                    icon.classList.replace('fa-chevron-up', 'fa-chevron-down');
                }
            }
        },

        saveEmailAutomationSettings: async function() {
            const isAutoPilot = document.getElementById('inpAutoPilot').checked;
            const onlyNew = document.getElementById('selEmailOnlyNew').value === 'true';
            const delayMinutes = parseInt(document.getElementById('inpEmailDelay').value) || 0;
            const scheduleTime = document.getElementById('inpEmailTime').value;

            this.emailConfig.autoPilot = isAutoPilot;
            this.emailConfig.onlyNew = onlyNew;
            this.emailConfig.delayMinutes = delayMinutes;
            this.emailConfig.scheduleTime = scheduleTime;

            try {
                await setDoc(doc(db, 'artifacts', appId, 'public', 'data', 'song_siang_settings', 'config'), { 
                    emailConfig: this.emailConfig
                }, { merge: true });

                if (isAutoPilot && !this.autoPilotInterval) this.startAutoPilot();
                if (!isAutoPilot && this.autoPilotInterval) this.stopAutoPilot();

                this.isAutoPilotExpanded = false;
                const content = document.getElementById('autoPilotSettingsContent');
                const icon = document.getElementById('autoPilotCollapseIcon');
                if(content) content.classList.add('hidden');
                if(icon) icon.classList.replace('fa-chevron-up', 'fa-chevron-down');

                Swal.fire({
                    iconHtml: '<i class="fa-solid fa-check text-primary text-5xl"></i>',
                    title: 'บันทึกสำเร็จ!',
                    showConfirmButton: false,
                    timer: 1500
                });
            } catch(e) {
                Swal.fire('ข้อผิดพลาด', 'ไม่สามารถบันทึกการตั้งค่าได้', 'error');
            }
        },

        showAutoPilotTargetSelectModal: function() {
            const usersWithMessages = predefinedUsers.map(user => {
                const userMsgs = this.allMessages.filter(m => this.isMsgEligibleForUser(m, user.studentId));
                return { ...user, msgCount: userMsgs.length };
            }).filter(u => u.msgCount > 0);

            if (usersWithMessages.length === 0) return;

            let excludedList = this.emailConfig.excludedAutoPilotUsers || [];

            let html = `
                <div class="text-left mb-4">
                    <div class="font-bold text-gray-700 text-sm mb-1">เป้าหมายส่งอัตโนมัติ</div>
                    <div class="text-xs text-gray-500">ติ๊กเลือกผู้ที่ต้องการให้ระบบส่งอีเมลอัตโนมัติ</div>
                </div>
                <div class="flex justify-between items-center mb-3 px-1">
                    <span class="text-xs font-bold text-primary" id="autoPilotSelectCount">เลือกแล้ว ${usersWithMessages.length - excludedList.length} / ${usersWithMessages.length}</span>
                    <button type="button" onclick="app.toggleAutoPilotSelectAll()" class="text-xs bg-blue-100 text-blue-600 px-3 py-1 rounded-full font-bold hover:bg-blue-200 transition-colors shadow-sm" id="btnAutoPilotSelectAll">
                        ${usersWithMessages.length - excludedList.length === usersWithMessages.length ? 'ยกเลิกทั้งหมด' : 'เลือกทั้งหมด'}
                    </button>
                </div>
                <div class="max-h-[45vh] overflow-y-auto pr-2 custom-scrollbar text-left" id="autoPilotTargetList">
            `;

            usersWithMessages.forEach(u => {
                const isChecked = !excludedList.includes(u.studentId);
                html += `
                    <label class="relative flex items-center gap-3 p-3 bg-white border border-gray-200 rounded-xl cursor-pointer hover:bg-gray-50 transition-colors mb-2">
                        <input type="checkbox" class="autopilot-target-checkbox sr-only" value="${u.studentId}" onchange="app.updateAutoPilotSelectCount()" ${isChecked ? 'checked' : ''}>
                        <div class="w-5 h-5 rounded border-2 border-gray-300 bg-white flex items-center justify-center shrink-0 transition-colors">
                            <i class="fa-solid fa-check text-white text-[10px] check-icon ${isChecked ? 'block' : 'hidden'}"></i>
                        </div>
                        <div class="flex items-center gap-2 min-w-0 flex-1">
                            <div class="w-8 h-8 rounded-full overflow-hidden bg-secondary border border-gray-200 shrink-0 relative flex items-center justify-center">
                                ${this.renderImage(`./65/${u.id}.png`, u.name)}
                                <span class="fallback-name absolute inset-0 z-10 hidden flex-center-all bg-secondary text-primary font-bold text-[10px] h-full w-full">${this.getAbbr(u.name)}</span>
                            </div>
                            <div class="min-w-0">
                                <div class="font-bold text-primary text-sm truncate">${u.name}</div>
                                <div class="text-[10px] text-gray-500">${u.msgCount} ข้อความค้างส่ง</div>
                            </div>
                        </div>
                    </label>
                `;
            });
            
            html += `</div>
                <style>
                    .autopilot-target-checkbox:checked + div { border-color: #2b71b8; background-color: #2b71b8; }
                    .autopilot-target-checkbox:checked + div .check-icon { display: block !important; }
                </style>
                <button onclick="app.confirmAutoPilotTargets()" class="w-full mt-4 py-3 bg-secondary hover:bg-yellow-400 text-primary rounded-xl font-bold text-sm shadow-md transition-transform transform active:scale-95 flex justify-center items-center gap-2">
                    <i class="fa-solid fa-floppy-disk"></i> บันทึกเป้าหมายชั่วคราว
                </button>
                <div class="text-[10px] text-center text-gray-400 mt-2">*อย่าลืมกด "บันทึกตั้งค่าระบบส่งอัตโนมัติ" ในหน้าหลักอีกครั้ง</div>
            `;

            Swal.fire({
                html: html,
                showConfirmButton: false,
                showCloseButton: true,
                customClass: {
                    popup: 'rounded-[2rem] border border-white/50 shadow-2xl backdrop-blur-xl w-[90%] max-w-sm sm:max-w-md',
                }
            });
        },

        updateAutoPilotSelectCount: function() {
            const checkboxes = document.querySelectorAll('.autopilot-target-checkbox');
            const checked = document.querySelectorAll('.autopilot-target-checkbox:checked');
            const countDisplay = document.getElementById('autoPilotSelectCount');
            const btnSelectAll = document.getElementById('btnAutoPilotSelectAll');
            
            if (countDisplay) countDisplay.innerText = `เลือกแล้ว ${checked.length} / ${checkboxes.length}`;
            if (btnSelectAll) {
                btnSelectAll.innerText = checked.length === checkboxes.length ? 'ยกเลิกทั้งหมด' : 'เลือกทั้งหมด';
            }
        },

        toggleAutoPilotSelectAll: function() {
            const checkboxes = document.querySelectorAll('.autopilot-target-checkbox');
            const checked = document.querySelectorAll('.autopilot-target-checkbox:checked');
            const isAllChecked = checked.length === checkboxes.length;
            
            checkboxes.forEach(cb => {
                cb.checked = !isAllChecked;
            });
            this.updateAutoPilotSelectCount();
        },

        confirmAutoPilotTargets: function() {
            const checkboxes = document.querySelectorAll('.autopilot-target-checkbox');
            let excluded = [];
            checkboxes.forEach(cb => {
                if (!cb.checked) excluded.push(cb.value);
            });
            
            this.emailConfig.excludedAutoPilotUsers = excluded;
            
            const root = document.getElementById('appRoot');
            if (root && this.currentView === 'admin_email_system') {
                root.innerHTML = this.renderAdminEmailSystem();
            }
            Swal.close();
        },

        renderAdminEmailSystem: function() {
            // คัดเฉพาะคนที่มียอดข้อความ "ค้างส่ง" (อิงรายบุคคล)
            const usersWithMessages = predefinedUsers.map(user => {
                const userMsgs = this.allMessages.filter(m => this.isMsgEligibleForUser(m, user.studentId));
                return { ...user, msgCount: userMsgs.length, messages: userMsgs };
            }).filter(u => u.msgCount > 0).sort((a, b) => b.msgCount - a.msgCount);

            // คัดข้อความทั้งหมดที่แต่ละคนมีในระบบ (เพื่อให้แสดงในรายการด้านล่าง)
            let listHtml = '';
            const allUsersWithMsgsList = predefinedUsers.map(user => {
                const userMsgs = this.allMessages.filter(m => m.to && Array.isArray(m.to) && (m.to.includes(user.studentId) || m.to.length >= predefinedUsers.length));
                return { ...user, msgCount: userMsgs.length, messages: userMsgs };
            }).filter(u => u.msgCount > 0).sort((a, b) => b.msgCount - a.msgCount);

            if (allUsersWithMsgsList.length === 0) {
                listHtml = `<div class="text-center py-12 text-white/70 font-medium">ไม่มีผู้รับข้อความในระบบ</div>`;
            } else {
                allUsersWithMsgsList.forEach(u => {
                    const abbr = this.getAbbr(u.name);
                    
                    // หายอดที่ค้างส่งของคนๆ นี้ เพื่อนำมาแสดงในวงเล็บ
                    const pendingCount = usersWithMessages.find(pendingUser => pendingUser.studentId === u.studentId)?.msgCount || 0;
                    const pendingBadge = pendingCount > 0 ? `<span class="bg-red-500 text-white px-2 py-0.5 rounded-full text-[9px] ml-1 shadow-sm border border-red-400">ค้าง ${pendingCount}</span>` : '';

                    listHtml += `
                    <div class="bg-white/60 p-4 rounded-2xl flex items-center justify-between border border-white/40 shadow-sm hover:bg-white/80 transition-colors">
                        <div class="flex items-center gap-3">
                            <div class="w-10 h-10 rounded-full bg-secondary flex items-center justify-center border border-white overflow-hidden relative shadow-sm shrink-0">
                                ${this.renderImage(`./65/${u.id}.png`, u.name)}
                                <span class="fallback-name absolute inset-0 z-10 hidden flex-center-all bg-secondary text-primary font-bold text-sm h-full w-full">${abbr}</span>
                            </div>
                            <div>
                                <div class="font-extrabold text-primary text-sm">${u.name}</div>
                                <div class="text-[10px] text-primary/60">${u.studentId}</div>
                            </div>
                        </div>
                        <div class="flex items-center gap-3">
                            <div class="bg-secondary text-primary px-3 py-1 rounded-full text-xs font-bold border border-white shadow-sm">
                                ทั้งหมด ${u.msgCount} ${pendingBadge}
                            </div>
                            <button onclick="app.showEmailMessageSelectModal('${u.studentId}')" class="bg-primary hover:bg-blue-600 text-white w-9 h-9 rounded-full flex items-center justify-center shadow-md transition-transform transform active:scale-95">
                                <i class="fa-solid fa-paper-plane text-xs"></i>
                            </button>
                        </div>
                    </div>`;
                });
            }

            const daysList = [
                {val: '1', label: 'จันทร์'}, {val: '2', label: 'อังคาร'}, {val: '3', label: 'พุธ'},
                {val: '4', label: 'พฤหัสฯ'}, {val: '5', label: 'ศุกร์'}, {val: '6', label: 'เสาร์'}, {val: '0', label: 'อาทิตย์'}
            ];

            let daysHtml = daysList.map(d => {
                const isSelected = this.emailConfig.scheduleDays.includes(d.val);
                return `<button id="btnDay${d.val}" onclick="app.toggleScheduleDay('${d.val}')" class="px-3 py-1 rounded-lg text-xs font-bold transition-colors border ${isSelected ? 'bg-secondary text-primary border-secondary' : 'bg-white/10 text-white/60 border-white/30'}">${d.label}</button>`;
            }).join('');

            return `
            <div class="glass-card rounded-[2rem] p-4 sm:p-8 shadow-2xl max-w-4xl mx-auto min-h-[75vh] mb-20 animate-slide-up flex flex-col">
                <div class="flex items-center justify-between mb-6 border-b border-white/40 pb-5 flex-wrap gap-4">
                    <div class="flex items-center">
                        <button onclick="app.showView('home')" class="p-3 bg-white/20 hover:bg-white/40 rounded-2xl mr-4 transition text-white shadow-sm">
                            <i class="fa-solid fa-arrow-left text-xl"></i>
                        </button>
                        <div>
                            <h2 class="text-xl font-extrabold text-white drop-shadow-md">ส่งอีเมลแจ้งเตือน</h2>
                            <p class="text-blue-100 text-xs font-bold">ส่งข้อความแจ้งเตือนเข้าอีเมลสมาชิก</p>
                        </div>
                    </div>
                    <div class="flex flex-wrap items-center gap-2">
                        <button onclick="app.resetEmailStatus()" class="bg-gray-500 hover:bg-gray-600 text-white px-4 py-2.5 rounded-xl text-xs font-bold shadow-lg transition-all flex items-center gap-2">
                            <i class="fa-solid fa-rotate-left"></i> <span class="hidden sm:inline">รีเซ็ตสถานะการส่ง</span>
                        </button>
                        <button onclick="app.startBulkEmailProcess()" class="bg-green-500 hover:bg-green-600 text-white px-4 py-2.5 rounded-xl text-xs font-bold shadow-lg transition-all flex items-center gap-2">
                            <i class="fa-solid fa-mail-bulk"></i> <span class="hidden sm:inline">ส่งให้ทุกคนที่ได้รับข้อความ</span>
                        </button>
                    </div>
                </div>

                <div class="bg-blue-500/20 border border-blue-200/30 p-4 rounded-xl text-white text-xs mb-6 leading-relaxed">
                    <i class="fa-solid fa-circle-info mr-1 text-secondary"></i> ระบบจะจัดรูปแบบข้อความเป็นอีเมลที่สวยงาม ส่งตรงไปยังอีเมลสถาบันของสมาชิกที่เลือก (รวดเร็วขึ้นเนื่องจากไม่ส่งเป็นรูปภาพแล้ว)
                </div>

                <div class="bg-white/10 rounded-2xl border border-white/20 mb-6 shadow-inner overflow-hidden">
                    <div class="p-5 cursor-pointer flex justify-between items-center hover:bg-white/5 transition-colors" onclick="app.toggleAutoPilotCollapse()">
                        <h3 class="text-white font-bold text-lg flex items-center mb-0">
                            <i class="fa-solid fa-robot text-secondary mr-2"></i> ตั้งค่าส่งอีเมลอัตโนมัติ (Auto-Pilot)
                        </h3>
                        <i id="autoPilotCollapseIcon" class="fa-solid fa-chevron-${this.isAutoPilotExpanded ? 'up' : 'down'} text-white/70 transition-transform text-xl"></i>
                    </div>
                    
                    <div id="autoPilotSettingsContent" class="p-5 pt-0 border-t border-white/10 mt-2 ${this.isAutoPilotExpanded ? '' : 'hidden'}">
                        <div class="flex justify-between items-center mb-4 pb-4 border-b border-white/10 mt-4">
                            <div>
                                <div class="text-sm font-bold text-white">เปิดใช้งานระบบส่งอัตโนมัติ (รันบนเครื่องนี้)</div>
                                <div class="text-[10px] text-white/60">ให้ระบบตรวจสอบและส่งอีเมลอัตโนมัติ<span class="text-secondary font-bold">เฉพาะตอนที่เปิดหน้านี้ทิ้งไว้เท่านั้น</span></div>
                            </div>
                            <label class="relative inline-flex items-center cursor-pointer">
                                <input type="checkbox" id="inpAutoPilot" class="sr-only peer" ${this.emailConfig.autoPilot ? 'checked' : ''}>
                                <div class="w-11 h-6 bg-white/30 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-secondary border border-white/50 shadow-inner"></div>
                            </label>
                        </div>
                        
                        <div class="mb-4">
                            <label class="block text-[10px] font-bold text-white mb-2">เงื่อนไขข้อความ</label>
                            <div class="glass-input w-full px-4 py-2 rounded-xl font-bold text-primary text-xs bg-white/50 text-center shadow-inner">
                                <i class="fa-solid fa-shield-halved mr-1"></i> ส่งเฉพาะข้อความที่ยังไม่เคยส่ง (ป้องกันส่งซ้ำ)
                            </div>
                            <input type="hidden" id="selEmailOnlyNew" value="true">
                        </div>

                        <div class="mb-4">
                            <label class="block text-[10px] font-bold text-white mb-2">รูปแบบการส่ง</label>
                            <div class="flex gap-2 bg-black/20 p-1 rounded-xl">
                                <button onclick="app.setEmailSendType('delay')" class="flex-1 py-2 rounded-xl text-xs font-bold transition-colors border ${this.emailConfig.sendType === 'delay' ? 'bg-secondary text-primary border-secondary' : 'bg-transparent text-white border-white/40'}">ส่งเมื่อครบเวลา (Delay)</button>
                                <button onclick="app.setEmailSendType('schedule')" class="flex-1 py-2 rounded-xl text-xs font-bold transition-colors border ${this.emailConfig.sendType === 'schedule' ? 'bg-secondary text-primary border-secondary' : 'bg-transparent text-white border-white/40'}">ตั้งเวลาส่ง (Schedule)</button>
                            </div>
                        </div>

                        <div id="emailDelaySettings" class="${this.emailConfig.sendType === 'delay' ? 'block' : 'hidden'} mb-4 p-4 bg-black/10 rounded-xl border border-white/10">
                            <label class="block text-[10px] font-bold text-white mb-2">หน่วงเวลาส่งหลังจากมีข้อความใหม่ (นาที)</label>
                            <input type="number" id="inpEmailDelay" value="${this.emailConfig.delayMinutes}" min="0" class="glass-input w-full px-4 py-2 rounded-xl outline-none font-bold text-primary text-center">
                            <p class="text-[9px] text-white/50 mt-1">* 0 = ส่งทันทีที่ระบบเช็คเจอ (เช็คทุก 1 นาที)</p>
                        </div>

                        <div id="emailScheduleSettings" class="${this.emailConfig.sendType === 'schedule' ? 'block' : 'hidden'} mb-4 p-4 bg-black/10 rounded-xl border border-white/10">
                            <label class="block text-[10px] font-bold text-white mb-2">เวลาที่ส่ง (เช่น 18:00)</label>
                            <input type="time" id="inpEmailTime" value="${this.emailConfig.scheduleTime}" class="glass-input w-full px-4 py-2 rounded-xl outline-none font-bold text-primary text-center mb-3">
                            
                            <label class="block text-[10px] font-bold text-white mb-2">วันที่ต้องการส่ง</label>
                            <div class="flex flex-wrap gap-1">
                                ${daysHtml}
                            </div>
                        </div>

                        <div class="mb-4 p-4 bg-black/20 rounded-xl border border-white/10 shadow-inner">
                            <label class="block text-[11px] font-bold text-white mb-2"><i class="fa-solid fa-users-viewfinder text-secondary mr-1"></i> เป้าหมายที่จะส่งอัตโนมัติ (ตามเงื่อนไขด้านบน)</label>
                            <div class="text-xs text-white/90 leading-relaxed">
                                ${usersWithMessages.length > 0
                                    ? `<button type="button" id="autoPilotTargetBtn" onclick="app.showAutoPilotTargetSelectModal()" class="w-full mt-1 bg-white/10 hover:bg-white/20 border border-white/20 rounded-xl p-3 text-left transition-colors flex justify-between items-center group shadow-sm">
                                           <div>
                                               <div class="font-bold text-secondary text-sm mb-0.5"><i class="fa-solid fa-list-check mr-1"></i> กดเพื่อเลือกเป้าหมาย</div>
                                               <div id="autoPilotTargetCountText" class="text-[10px] text-white/70">ส่งให้ ${usersWithMessages.filter(u => !(this.emailConfig.excludedAutoPilotUsers || []).includes(u.studentId)).length} จาก ${usersWithMessages.length} ท่าน ที่มีข้อความค้างส่ง</div>
                                           </div>
                                           <i class="fa-solid fa-chevron-right text-white/50 group-hover:text-white transition-colors"></i>
                                       </button>`
                                    : '<span class="text-green-400 font-bold"><i class="fa-solid fa-check-circle mr-1"></i> ไม่มีข้อความใหม่ที่ค้างส่ง</span>'}
                            </div>
                        </div>

                        <button onclick="app.saveEmailAutomationSettings()" class="w-full py-3 bg-secondary hover:bg-yellow-400 text-primary rounded-xl font-extrabold text-sm shadow-lg transition-transform hover:-translate-y-1 mt-2">
                            <i class="fa-solid fa-floppy-disk mr-2"></i> บันทึกตั้งค่าระบบส่งอัตโนมัติ (รันบนเครื่องนี้)
                        </button>
                    </div>
                </div>

                <h3 class="text-white font-bold text-sm mb-3 pl-2"><i class="fa-solid fa-users mr-2 text-secondary"></i>รายชื่อสมาชิกที่ได้รับข้อความ (${allUsersWithMsgsList.length} ท่าน)</h3>
                <div class="flex-1 overflow-y-auto pr-1 space-y-3 custom-scrollbar">
                    ${listHtml}
                </div>
            </div>`;
        },

        resetEmailStatus: function() {
            const msgsToReset = this.allMessages.filter(m => m.emailSent);
            if (msgsToReset.length === 0) return Swal.fire('แจ้งเตือน', 'ไม่มีข้อความที่เคยส่งอีเมลแล้วให้รีเซ็ต', 'info');

            const resetTargets = new Set();
            msgsToReset.forEach(msg => {
                if (msg.to && Array.isArray(msg.to)) {
                    if (msg.to.length >= predefinedUsers.length) {
                        predefinedUsers.forEach(u => resetTargets.add(u.studentId));
                    } else {
                        msg.to.forEach(uid => resetTargets.add(uid));
                    }
                }
            });

            const usersWithResettableMsgs = predefinedUsers.filter(u => resetTargets.has(u.studentId));

            if (usersWithResettableMsgs.length === 0) return;

            let html = `
                <div class="text-left mb-4">
                    <div class="font-bold text-gray-700 text-sm mb-1">เลือกผู้ที่ต้องการรีเซ็ตสถานะ</div>
                    <div class="text-xs text-gray-500">ระบบจะรีเซ็ตเฉพาะคนที่คุณเลือกเท่านั้น <br><span class="text-[10px] text-green-600">*คนอื่นๆ ในกลุ่มข้อความเดียวกันจะไม่ได้รับอีเมลซ้ำ</span></div>
                </div>
                <div class="flex justify-between items-center mb-3 px-1">
                    <span class="text-xs font-bold text-primary" id="resetSelectCount">เลือกแล้ว 0 / ${usersWithResettableMsgs.length}</span>
                    <button type="button" onclick="app.toggleResetSelectAll()" class="text-xs bg-blue-100 text-blue-600 px-3 py-1 rounded-full font-bold hover:bg-blue-200 transition-colors shadow-sm" id="btnResetSelectAll">
                        เลือกทั้งหมด
                    </button>
                </div>
                <div class="max-h-[45vh] overflow-y-auto pr-2 custom-scrollbar text-left" id="resetTargetList">
            `;

            usersWithResettableMsgs.forEach(u => {
                const userSentMsgs = msgsToReset.filter(m => m.to && (m.to.includes(u.studentId) || m.to.length >= predefinedUsers.length));
                
                html += `
                    <label class="relative flex items-center gap-3 p-3 bg-white border border-gray-200 rounded-xl cursor-pointer hover:bg-gray-50 transition-colors mb-2">
                        <input type="checkbox" class="reset-target-checkbox sr-only" value="${u.studentId}" onchange="app.updateResetSelectCount()">
                        <div class="w-5 h-5 rounded border-2 border-gray-300 bg-white flex items-center justify-center shrink-0 transition-colors">
                            <i class="fa-solid fa-check text-white text-[10px] check-icon hidden"></i>
                        </div>
                        <div class="flex items-center gap-2 min-w-0 flex-1">
                            <div class="w-8 h-8 rounded-full overflow-hidden bg-secondary border border-gray-200 shrink-0 relative flex items-center justify-center">
                                ${this.renderImage(`./65/${u.id}.png`, u.name)}
                                <span class="fallback-name absolute inset-0 z-10 hidden flex-center-all bg-secondary text-primary font-bold text-[10px] h-full w-full">${this.getAbbr(u.name)}</span>
                            </div>
                            <div class="min-w-0">
                                <div class="font-bold text-primary text-sm truncate">${u.name}</div>
                                <div class="text-[10px] text-gray-500">มี ${userSentMsgs.length} ข้อความที่รีเซ็ตได้</div>
                            </div>
                        </div>
                    </label>
                `;
            });
            
            html += `</div>
                <style>
                    .reset-target-checkbox:checked + div { border-color: #2b71b8; background-color: #2b71b8; }
                    .reset-target-checkbox:checked + div .check-icon { display: block !important; }
                </style>
                <button onclick="app.confirmResetEmailStatus()" class="w-full mt-4 py-3 bg-secondary hover:bg-yellow-400 text-primary rounded-xl font-bold text-sm shadow-md transition-transform transform active:scale-95 flex justify-center items-center gap-2">
                    <i class="fa-solid fa-rotate-left"></i> ยืนยันการรีเซ็ต
                </button>
            `;

            Swal.fire({
                html: html,
                showConfirmButton: false,
                showCloseButton: true,
                customClass: {
                    popup: 'rounded-[2rem] border border-white/50 shadow-2xl backdrop-blur-xl w-[90%] max-w-sm sm:max-w-md',
                }
            });
        },

        updateResetSelectCount: function() {
            const checkboxes = document.querySelectorAll('.reset-target-checkbox');
            const checked = document.querySelectorAll('.reset-target-checkbox:checked');
            const countDisplay = document.getElementById('resetSelectCount');
            const btnSelectAll = document.getElementById('btnResetSelectAll');
            
            if (countDisplay) countDisplay.innerText = `เลือกแล้ว ${checked.length} / ${checkboxes.length}`;
            if (btnSelectAll) {
                btnSelectAll.innerText = checked.length === checkboxes.length ? 'ยกเลิกทั้งหมด' : 'เลือกทั้งหมด';
            }
        },

        toggleResetSelectAll: function() {
            const checkboxes = document.querySelectorAll('.reset-target-checkbox');
            const checked = document.querySelectorAll('.reset-target-checkbox:checked');
            const isAllChecked = checked.length === checkboxes.length;
            
            checkboxes.forEach(cb => {
                cb.checked = !isAllChecked;
            });
            this.updateResetSelectCount();
        },

        confirmResetEmailStatus: async function() {
            const checkboxes = document.querySelectorAll('.reset-target-checkbox:checked');
            if (checkboxes.length === 0) return Swal.fire('แจ้งเตือน', 'กรุณาเลือกผู้ใช้อย่างน้อย 1 คน', 'warning');
            
            const selectedUids = Array.from(checkboxes).map(cb => cb.value);

            Swal.fire({
                title: 'กำลังรีเซ็ต...',
                html: '<i class="fa-solid fa-circle-notch fa-spin text-4xl text-primary mt-4 mb-3"></i>',
                showConfirmButton: false,
                allowOutsideClick: false
            });

            try {
                const batch = writeBatch(db);
                const msgsToReset = this.allMessages.filter(m => m.emailSent);
                let resetCount = 0;

                msgsToReset.forEach(m => {
                    const targetUids = selectedUids.filter(uid => m.to && (m.to.includes(uid) || m.to.length >= predefinedUsers.length));
                    
                    if (targetUids.length > 0) {
                        const ref = doc(db, 'artifacts', appId, 'public', 'data', 'song_siang_messages', m.id);
                        let currentReset = m.resetForUsers ? [...m.resetForUsers] : [];
                        let modified = false;
                        
                        targetUids.forEach(uid => {
                            if (!currentReset.includes(uid)) {
                                currentReset.push(uid);
                                modified = true;
                            }
                        });
                        
                        // อัปเดตเฉพาะอาเรย์ resetForUsers โดยไม่ต้องปรับ emailSent หลัก
                        if (modified) {
                            batch.update(ref, { resetForUsers: currentReset }); 
                            resetCount++;
                        }
                    }
                });

                if (resetCount > 0) {
                    await batch.commit();
                    Swal.fire({
                        iconHtml: '<i class="fa-solid fa-check-circle text-green-500 text-5xl"></i>',
                        title: 'รีเซ็ตสำเร็จ!',
                        text: `ดึงข้อความกลับมาค้างส่งให้คนที่เลือกเรียบร้อยแล้ว`,
                        showConfirmButton: false,
                        timer: 2000
                    });
                    this.showView('admin_email_system');
                } else {
                    Swal.fire('สำเร็จ', 'ไม่มีข้อความให้รีเซ็ตสำหรับคนที่เลือก', 'info');
                }
            } catch (e) {
                Swal.fire('ข้อผิดพลาด', 'ไม่สามารถรีเซ็ตได้', 'error');
                console.error(e);
            }
        },

        showEmailMessageSelectModal: function(uid) {
            const user = predefinedUsers.find(u => u.studentId === uid);
            if (!user) return;

            // คัดเฉพาะข้อความที่อยู่ในคิวค้างส่ง (อิงเงื่อนไขรายบุคคล)
            const userMsgs = this.allMessages.filter(m => this.isMsgEligibleForUser(m, uid));
            
            if (userMsgs.length === 0) return Swal.fire('แจ้งเตือน', 'ไม่มีข้อความค้างส่งสำหรับสมาชิกท่านนี้', 'info');

            let msgsHtml = '';
            userMsgs.forEach(msg => {
                const isAudio = msg.type === 'audio';
                const isAnonymous = msg.from.isAnonymous;
                let displayName = isAnonymous ? 'ผู้ไม่ประสงค์ออกนาม' : msg.from.name;
                const dateStr = new Date(msg.timestamp).toLocaleString('th-TH');
                
                msgsHtml += `
                    <label class="relative flex items-start gap-3 p-3 bg-white border border-gray-200 rounded-xl cursor-pointer hover:bg-gray-50 transition-colors mb-2">
                        <input type="checkbox" class="msg-checkbox sr-only" value="${msg.id}" checked>
                        <div class="w-5 h-5 rounded border-2 border-gray-300 bg-white flex items-center justify-center shrink-0 mt-0.5 transition-colors">
                            <i class="fa-solid fa-check text-white text-[10px] check-icon hidden"></i>
                        </div>
                        <div class="flex-1 min-w-0 text-left">
                            <div class="flex justify-between items-center mb-1">
                                <span class="font-bold text-primary text-xs truncate">${displayName}</span>
                                <span class="text-[9px] text-gray-500">${dateStr}</span>
                            </div>
                            <div class="text-xs text-gray-700 truncate">
                                ${isAudio ? `<i class="fa-solid fa-microphone text-secondary mr-1"></i> คลิปเสียง (${msg.audioDuration}วิ)` : msg.content}
                            </div>
                        </div>
                    </label>
                `;
            });

            const contentHtml = `
                <div class="text-left mb-4">
                    <div class="font-bold text-gray-700 text-sm mb-1">เลือกข้อความที่จะส่งให้ ${user.name}</div>
                    <div class="text-xs text-gray-500">ติ๊กเลือกข้อความที่ต้องการส่ง (เฉพาะที่ค้างส่งอยู่)</div>
                </div>
                <div class="max-h-[50vh] overflow-y-auto pr-2 custom-scrollbar" id="emailMsgSelectContainer">
                    ${msgsHtml}
                </div>
                <button onclick="app.startEmailProcess('${uid}')" class="w-full mt-4 py-3 bg-secondary hover:bg-yellow-400 text-primary rounded-xl font-bold text-sm shadow-md transition-transform transform active:scale-95 flex justify-center items-center gap-2">
                    <i class="fa-solid fa-paper-plane"></i> ยืนยันและส่งอีเมล
                </button>
            `;

            Swal.fire({
                title: 'เลือกข้อความ',
                html: contentHtml,
                showConfirmButton: false,
                showCloseButton: true,
                customClass: {
                    popup: 'rounded-[2rem] border border-white/50 shadow-2xl backdrop-blur-xl w-[90%] max-w-sm sm:max-w-md',
                    title: 'text-primary font-extrabold text-xl md:text-2xl',
                }
            });
        },

        buildEmailTemplate: function(user, msgs) {
            const todayStr = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
            let msgsListHtml = '';
            
            const baseUrl = "https://iearchitecture65.github.io/Song-Siang/";

            const parseTextWithEmojis = (text) => {
                if (!text) return '';
                let htmlText = text.replace(/\n/g, '<br>');
                
                return htmlText.replace(/[\p{Emoji_Presentation}\p{Extended_Pictographic}]/gu, match => {
                    const hexPoints = [];
                    for (let i = 0; i < match.length; i++) {
                        const hex = match.codePointAt(i).toString(16);
                        if (match.codePointAt(i) > 0xFFFF) i++; 
                        if (hex !== 'fe0f') hexPoints.push(hex); 
                    }
                    const twemojiCode = hexPoints.join('-');
                    return `<img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/72x72/${twemojiCode}.png" alt="${match}" style="width: 20px; height: 20px; vertical-align: middle; display: inline-block; margin: 0 2px; border: none; box-shadow: none;" />`;
                });
            };

            msgs.forEach((msg, idx) => {
                const dateStr = new Date(msg.timestamp).toLocaleString('th-TH');
                const isAudio = msg.type === 'audio';
                let displayName = msg.from.isAnonymous ? 'ผู้ไม่ประสงค์ออกนาม' : msg.from.name;
                const abbr = this.getAbbr(msg.from.name);
                
                let msgContentHtml = '';
                if (isAudio) {
                    const link = `${baseUrl}?m=${msg.id}`;
                    msgContentHtml = `
                        <div style="background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 12px; padding: 12px 16px; display: inline-block;">
                            <div style="color: #2b71b8; font-weight: bold; font-size: 14px; margin-bottom: 8px;">
                                <img src="https://cdn-icons-png.flaticon.com/512/25/25694.png" width="14" style="vertical-align: middle; opacity: 0.7; margin-right: 5px;"> 
                                คลิปเสียง (${msg.audioDuration || '?'} วินาที)
                            </div>
                            <a href="${link}" style="display: inline-block; background-color: #2b71b8; color: #ffffff; text-decoration: none; padding: 8px 20px; border-radius: 10px; font-size: 13px; font-weight: bold;">
                                <img src="https://cdn-icons-png.flaticon.com/512/724/724930.png" width="12" style="vertical-align: middle; margin-right: 5px; filter: brightness(0) invert(1);"> ฟังข้อความเสียง
                            </a>
                        </div>
                    `;
                } else {
                    const parsedContent = parseTextWithEmojis(msg.content.trim());
                    msgContentHtml = `
                        <div style="background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 16px; padding: 16px; color: #1e293b; font-size: 15px; font-weight: 500; line-height: 1.6; position: relative;">
                            <div style="font-size: 28px; color: #cbd5e1; position: absolute; top: 8px; left: 12px; font-family: Georgia, serif; line-height: 1;">"</div>
                            <div style="padding-left: 20px; word-break: break-word;">${parsedContent}</div>
                        </div>
                    `;
                }

                msgsListHtml += `
                    <div style="background-color: #ffffff; border: 2px solid #e2e8f0; border-radius: 20px; padding: 20px; margin-bottom: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.03);">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 14px;">
                            <tr>
                                <td width="55" valign="top">
                                    <div style="width: 48px; height: 48px; background-color: #fddf31; border-radius: 50%; text-align: center; line-height: 48px; font-size: 18px; font-weight: bold; color: #2b71b8; border: 2px solid #e2e8f0; overflow: hidden;">
                                        ${msg.from.isAnonymous ? '<img src="https://cdn-icons-png.flaticon.com/512/2202/2202112.png" width="24" style="vertical-align: middle; margin-top: 10px;">' : abbr}
                                    </div>
                                </td>
                                <td valign="middle" style="padding-left: 12px;">
                                    <div style="color: #2b71b8; font-weight: 800; font-size: 16px;">${displayName}</div>
                                    <div style="color: #64748b; font-size: 12px; margin-top: 3px;">
                                        <img src="https://cdn-icons-png.flaticon.com/512/2088/2088617.png" width="12" style="vertical-align: middle; margin-right: 4px; opacity: 0.6;">${dateStr}
                                    </div>
                                </td>
                            </tr>
                        </table>
                        ${msgContentHtml}
                    </div>
                `;
            });

            return `
            <!DOCTYPE html>
            <html lang="th">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>ข้อความใหม่จาก Song Siang</title>
                <style>
                    body { font-family: 'Prompt', Tahoma, sans-serif; background-color: #f1f5f9; margin: 0; padding: 0; }
                    .container { max-width: 600px; margin: 30px auto; background-color: #ffffff; border-radius: 20px; overflow: hidden; box-shadow: 0 10px 25px rgba(0,0,0,0.05); }
                    .header { background: linear-gradient(135deg, #1e5288 0%, #2b71b8 100%); padding: 30px 20px; text-align: center; color: white; }
                    .content { padding: 25px; background-color: #f1f5f9; }
                    .footer { background-color: #e2e8f0; padding: 20px; text-align: center; font-size: 12px; color: #64748b; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1 style="margin: 0; font-size: 26px; text-shadow: 0 2px 4px rgba(0,0,0,0.2);">ส่งเสียง <span style="color: #fddf31;">Song Siang</span></h1>
                        <p style="margin: 8px 0 0 0; opacity: 0.9; font-size: 14px; letter-spacing: 0.5px;">แพลตฟอร์มส่งต่อความรู้สึก</p>
                    </div>
                    <div class="content">
                        <div style="background-color: #ffffff; border-radius: 16px; padding: 20px; margin-bottom: 20px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
                            <h2 style="color: #1e293b; font-size: 20px; margin-top: 0; margin-bottom: 8px;">สวัสดี, ${user.name} <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/72x72/1f44b.png" width="20" style="vertical-align: middle; margin-left: 4px;"></h2>
                            <p style="color: #475569; font-size: 14px; margin: 0;">คุณมี <b>${msgs.length}</b> ข้อความใหม่ ประจำวันที่ ${todayStr}</p>
                        </div>
                        
                        <div>
                            ${msgsListHtml}
                        </div>
                        
                        <div style="text-align: center; margin-top: 30px; margin-bottom: 10px;">
                            <a href="${baseUrl}" style="display: inline-block; background-color: #fddf31; color: #2b71b8; text-decoration: none; padding: 14px 32px; border-radius: 25px; font-weight: 800; font-size: 16px; box-shadow: 0 4px 12px rgba(253,223,49,0.4);">
                                <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/72x72/1f449.png" width="18" style="vertical-align: middle; margin-right: 6px;"> เข้าสู่ระบบเพื่ออ่านและตอบกลับ
                            </a>
                        </div>
                    </div>
                    <div class="footer">
                        &copy; ${new Date().getFullYear()} ส่งเสียง (Song Siang). All rights reserved.<br>
                        หากมีข้อสงสัย หรือพบปัญหาการใช้งาน กรุณาติดต่อผู้ดูแลระบบ
                    </div>
                </div>
            </body>
            </html>
            `;
        },

        startEmailProcess: async function(uid) {
            const checkboxes = document.querySelectorAll('.msg-checkbox:checked');
            if (checkboxes.length === 0) return Swal.fire('แจ้งเตือน', 'กรุณาเลือกอย่างน้อย 1 ข้อความ', 'warning');
            
            const selectedMsgIds = Array.from(checkboxes).map(cb => cb.value);
            const msgsToSend = this.allMessages.filter(m => selectedMsgIds.includes(m.id));
            const user = predefinedUsers.find(u => u.studentId === uid);
            
            if (!user) return Swal.fire('เกิดข้อผิดพลาด', 'ไม่พบข้อมูลผู้ใช้', 'error');

            Swal.fire({
                title: 'กำลังส่งอีเมล...',
                html: '<div class="flex flex-col items-center"><i class="fa-solid fa-circle-notch fa-spin text-4xl text-primary mt-4 mb-3"></i><p class="text-sm text-gray-600 font-bold">กำลังจัดเตรียมข้อความ...</p></div>',
                showConfirmButton: false,
                allowOutsideClick: false
            });

            const htmlBody = this.buildEmailTemplate(user, msgsToSend);
            const email = `${uid}@kmitl.ac.th`;

            try {
                const response = await fetch(this.EMAIL_API_URL, {
                    method: "POST",
                    headers: { "Content-Type": "text/plain;charset=utf-8" },
                    body: JSON.stringify({
                        email: email,
                        subject: `คุณมี ${msgsToSend.length} ข้อความใหม่จาก ส่งเสียง`,
                        htmlBody: htmlBody
                    })
                });
                
                const resJson = await response.json();
                
                if (resJson.status === 'success') {
                    const batch = writeBatch(db);
                    msgsToSend.forEach(m => {
                        const ref = doc(db, 'artifacts', appId, 'public', 'data', 'song_siang_messages', m.id);
                        // ลบ uid ของคนที่ส่งสำเร็จออกจากคิวรีเซ็ต และอัปเดต emailSent เป็น true
                        let newReset = (m.resetForUsers || []).filter(id => id !== uid);
                        batch.update(ref, { emailSent: true, resetForUsers: newReset });
                    });
                    await batch.commit();

                    Swal.fire({
                        iconHtml: '<i class="fa-solid fa-paper-plane text-primary text-5xl"></i>',
                        title: 'ส่งอีเมลสำเร็จ!',
                        text: `ส่งไปยัง ${email} เรียบร้อยแล้ว`,
                        showConfirmButton: false,
                        timer: 2000
                    });
                    
                    this.showView('admin_email_system');
                } else {
                    Swal.fire('ข้อผิดพลาด', resJson.message || 'ไม่สามารถส่งอีเมลได้', 'error');
                }
            } catch (err) {
                console.error("Email API Error:", err);
                Swal.fire('ข้อผิดพลาด', 'เกิดปัญหาในการเชื่อมต่อกับเซิร์ฟเวอร์ส่งอีเมล', 'error');
            }
        },

        startBulkEmailProcess: async function() {
            // คัดเฉพาะคนที่มียอดข้อความ "ค้างส่ง"
            const usersWithMessages = predefinedUsers.map(user => {
                const userMsgs = this.allMessages.filter(m => this.isMsgEligibleForUser(m, user.studentId));
                return { ...user, messages: userMsgs };
            }).filter(u => u.messages.length > 0);

            if (usersWithMessages.length === 0) {
                return Swal.fire('แจ้งเตือน', 'ไม่มีข้อความใหม่ที่ยังไม่ได้ส่งอีเมล (ตามเงื่อนไขที่ตั้งค่าไว้)', 'warning');
            }

            Swal.fire({
                title: `ยืนยันการส่ง Bulk Email`,
                html: `ระบบจะทำการส่งอีเมลแจ้งเตือนไปยังสมาชิกที่ค้างส่งทั้งหมด <b>${usersWithMessages.length} ท่าน</b><br><br><span class="text-xs text-red-500">อาจใช้เวลาสักครู่ กรุณาอย่าปิดหน้าต่างนี้จนกว่าจะเสร็จสิ้น</span>`,
                icon: 'warning',
                showCancelButton: true,
                confirmButtonText: 'เริ่มส่งทันที',
                cancelButtonText: 'ยกเลิก'
            }).then(async (result) => {
                if (result.isConfirmed) {
                    let successCount = 0;
                    let failCount = 0;

                    Swal.fire({
                        title: 'กำลังประมวลผล...',
                        html: `
                            <div class="mb-4">
                                <i class="fa-solid fa-circle-notch fa-spin text-4xl text-primary mt-2 mb-3"></i>
                                <p class="font-bold text-gray-700 text-sm">กำลังส่งอีเมลทีละรายการ...</p>
                            </div>
                            <div class="flex justify-around bg-gray-50 p-3 rounded-xl border border-gray-200">
                                <div class="text-center">
                                    <div class="text-xs text-gray-500">ทั้งหมด</div>
                                    <div class="font-bold text-lg">${usersWithMessages.length}</div>
                                </div>
                                <div class="text-center">
                                    <div class="text-xs text-green-500">สำเร็จ</div>
                                    <div class="font-bold text-lg text-green-600" id="bulkSuccess">0</div>
                                </div>
                                <div class="text-center">
                                    <div class="text-xs text-red-500">ล้มเหลว</div>
                                    <div class="font-bold text-lg text-red-600" id="bulkFail">0</div>
                                </div>
                            </div>
                        `,
                        showConfirmButton: false,
                        allowOutsideClick: false
                    });

                    for (const user of usersWithMessages) {
                        const htmlBody = this.buildEmailTemplate(user, user.messages);
                        const email = `${user.studentId}@kmitl.ac.th`;

                        try {
                            const response = await fetch(this.EMAIL_API_URL, {
                                method: "POST",
                                headers: { "Content-Type": "text/plain;charset=utf-8" },
                                body: JSON.stringify({
                                    email: email,
                                    subject: `คุณมี ${user.messages.length} ข้อความใหม่จาก ส่งเสียง`,
                                    htmlBody: htmlBody
                                })
                            });
                            const resJson = await response.json();
                            
                            if (resJson.status === 'success') {
                                const batch = writeBatch(db);
                                user.messages.forEach(m => {
                                    const ref = doc(db, 'artifacts', appId, 'public', 'data', 'song_siang_messages', m.id);
                                    let newReset = (m.resetForUsers || []).filter(id => id !== user.studentId);
                                    batch.update(ref, { emailSent: true, resetForUsers: newReset });
                                });
                                await batch.commit();

                                successCount++;
                                document.getElementById('bulkSuccess').innerText = successCount;
                            } else {
                                failCount++;
                                document.getElementById('bulkFail').innerText = failCount;
                            }
                        } catch (err) {
                            failCount++;
                            document.getElementById('bulkFail').innerText = failCount;
                        }
                        await new Promise(r => setTimeout(r, 1500));
                    }

                    Swal.fire({
                        iconHtml: '<i class="fa-solid fa-flag-checkered text-primary text-5xl"></i>',
                        title: 'ส่งอีเมลเสร็จสิ้น!',
                        html: `สำเร็จ: <b class="text-green-500">${successCount}</b><br>ล้มเหลว: <b class="text-red-500">${failCount}</b>`,
                        confirmButtonText: 'รับทราบ'
                    }).then(() => {
                        this.showView('admin_email_system');
                    });
                }
            });
        }
    });
}
