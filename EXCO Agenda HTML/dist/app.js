const sections = [...document.querySelectorAll('.agenda-section')];
const navLinks = [...document.querySelectorAll('.agenda-nav a')];
const details = [...document.querySelectorAll('details')];
const progressBar = document.querySelector('#progressBar');
const expandButton = document.querySelector('#expandAll');
const presentButton = document.querySelector('#presentMode');
const printButton = document.querySelector('#printAgenda');
const editButton = document.querySelector('#editText');
const editBar = document.querySelector('#editBar');
const editStatus = document.querySelector('#editStatus');
const saveTextButton = document.querySelector('#saveText');
const formatButtons = [...document.querySelectorAll('[data-command]')];
const secretarialMessage = document.querySelector('#secretarialMessage');
const copySecretarialMessageButton = document.querySelector('#copySecretarialMessage');
const secretarialCopyStatus = document.querySelector('#secretarialCopyStatus');

const editableSelector = [
  '#agenda h1', '#agenda h2', '#agenda h3', '#agenda p', '#agenda li', '#agenda summary',
  '#agenda figcaption', '#agenda .meeting-objective > strong', '#agenda .hero-mark > span',
  '#agenda .decision-strip strong', '#agenda .decision-strip span',
  '#agenda .section-number', '#agenda .metric', '#agenda .model-label',
  '#agenda .sequence > section > span', '#agenda .eligibility > strong',
  '#agenda .eligibility > span', '#agenda .issue-prompts > section > span',
  '#agenda .merch-tiers > section > span'
].join(',');
const editableContent = [...document.querySelectorAll(editableSelector)]
  .filter(element => !element.closest('.action-register table'));
const contentStorageKey = 'utar-yoga-exco-editable-content-v2';
let saveTimer;

const sectionEditCounters = new Map();
editableContent.forEach(element => {
  const sectionId = element.closest('.agenda-section')?.id || 'agenda';
  const sectionIndex = sectionEditCounters.get(sectionId) || 0;
  sectionEditCounters.set(sectionId, sectionIndex + 1);
  element.dataset.editKey = `${sectionId}:${sectionIndex}`;
  element.spellcheck = true;
});

const readSavedContent = () => {
  try {
    const value = JSON.parse(localStorage.getItem(contentStorageKey));
    return value && typeof value === 'object' ? value : {};
  } catch {
    return {};
  }
};

const savedContent = readSavedContent();
editableContent.forEach(element => {
  const saved = savedContent[element.dataset.editKey];
  if (typeof saved === 'string' && !element.hasAttribute('data-edit-reset')) element.innerHTML = saved;
});

const classifyLongTitles = () => {
  document.querySelectorAll('.section-heading h2').forEach(title => {
    const length = title.textContent.trim().length;
    title.classList.toggle('long-title', length > 34 && length <= 50);
    title.classList.toggle('very-long-title', length > 50);
  });
};

const saveEditableContent = () => {
  const content = {};
  editableContent.forEach(element => { content[element.dataset.editKey] = element.innerHTML; });
  localStorage.setItem(contentStorageKey, JSON.stringify(content));
  editStatus.textContent = 'Changes saved';
};

const scheduleContentSave = () => {
  clearTimeout(saveTimer);
  editStatus.textContent = 'Saving…';
  saveTimer = setTimeout(saveEditableContent, 500);
};

editableContent.forEach(element => element.addEventListener('input', () => {
  classifyLongTitles();
  scheduleContentSave();
}));

const setEditing = enabled => {
  document.body.classList.toggle('editing', enabled);
  editableContent.forEach(element => { element.contentEditable = enabled ? 'true' : 'false'; });
  editButton.setAttribute('aria-pressed', String(enabled));
  editButton.textContent = enabled ? 'Finish editing' : 'Edit text';
  editBar.hidden = !enabled;
  editStatus.textContent = enabled ? 'Select text to format' : 'Changes saved';
  if (!enabled) saveEditableContent();
};

editButton.addEventListener('click', () => setEditing(!document.body.classList.contains('editing')));
saveTextButton.addEventListener('click', saveEditableContent);
formatButtons.forEach(button => {
  button.addEventListener('mousedown', event => event.preventDefault());
  button.addEventListener('click', () => {
    const selection = window.getSelection();
    const anchor = selection?.anchorNode?.parentElement;
    if (!selection || selection.isCollapsed || !anchor?.closest('[contenteditable="true"]')) {
      editStatus.textContent = 'Select agenda text first';
      return;
    }
    const command = button.dataset.command;
    const value = button.dataset.value || null;
    const applied = document.execCommand(command, false, value);
    if (!applied && command === 'hiliteColor') document.execCommand('backColor', false, value);
    scheduleContentSave();
  });
});

classifyLongTitles();

const secretarialMessageStorageKey = 'utar-yoga-secretarial-whatsapp-message-v2';
const savedSecretarialMessage = localStorage.getItem(secretarialMessageStorageKey);
if (savedSecretarialMessage) secretarialMessage.value = savedSecretarialMessage;
secretarialMessage.addEventListener('input', () => {
  localStorage.setItem(secretarialMessageStorageKey, secretarialMessage.value);
  secretarialCopyStatus.textContent = 'Message saved in this browser.';
});
copySecretarialMessageButton.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(secretarialMessage.value);
    secretarialCopyStatus.textContent = 'Copied — ready to paste into WhatsApp.';
  } catch {
    secretarialMessage.focus();
    secretarialMessage.select();
    document.execCommand('copy');
    secretarialCopyStatus.textContent = 'Copied — ready to paste into WhatsApp.';
  }
});

expandButton.textContent = details.some(item => !item.open) ? 'Expand all detail' : 'Collapse detail';

const updateProgress = () => {
  const scrollable = document.documentElement.scrollHeight - window.innerHeight;
  const ratio = scrollable > 0 ? window.scrollY / scrollable : 0;
  progressBar.style.width = `${Math.min(100, Math.max(0, ratio * 100))}%`;
};

const observer = new IntersectionObserver((entries) => {
  const visible = entries
    .filter(entry => entry.isIntersecting)
    .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
  if (!visible) return;
  navLinks.forEach(link => link.classList.toggle('active', link.hash === `#${visible.target.id}`));
}, { rootMargin: '-20% 0px -65%', threshold: [0, .15, .5] });

sections.forEach(section => observer.observe(section));
window.addEventListener('scroll', updateProgress, { passive: true });
updateProgress();

expandButton.addEventListener('click', () => {
  const shouldOpen = details.some(item => !item.open);
  details.forEach(item => { item.open = shouldOpen; });
  expandButton.textContent = shouldOpen ? 'Collapse detail' : 'Expand all detail';
});

presentButton.addEventListener('click', () => {
  document.body.classList.toggle('focus-mode');
  const enabled = document.body.classList.contains('focus-mode');
  presentButton.textContent = enabled ? 'Exit focus' : 'Focus mode';
  presentButton.setAttribute('aria-pressed', String(enabled));
});

printButton.addEventListener('click', () => window.print());

document.addEventListener('keydown', event => {
  if (!document.body.classList.contains('focus-mode')) return;
  if (!['ArrowDown', 'ArrowRight', 'PageDown', 'ArrowUp', 'ArrowLeft', 'PageUp'].includes(event.key)) return;
  if (['INPUT', 'SELECT', 'TEXTAREA'].includes(document.activeElement?.tagName)) return;
  event.preventDefault();
  const currentIndex = sections.findIndex(section => {
    const rect = section.getBoundingClientRect();
    return rect.top <= window.innerHeight * .35 && rect.bottom > window.innerHeight * .35;
  });
  const forward = ['ArrowDown', 'ArrowRight', 'PageDown'].includes(event.key);
  const nextIndex = Math.min(sections.length - 1, Math.max(0, currentIndex + (forward ? 1 : -1)));
  sections[nextIndex].scrollIntoView({ behavior: 'smooth', block: 'start' });
});

const statusOptions = ['Proposed', 'Not started', 'Drafting', 'In progress', 'Internal review', 'Awaiting approval', 'Approved', 'Scheduled', 'Published', 'Filed', 'Complete', 'Deferred'];

const cloneRows = rows => rows.map(row => ({ ...row }));

const getLocalDateValue = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const closedStatuses = new Set(['Approved', 'Published', 'Filed', 'Complete', 'Deferred']);

const updateRowDateState = (row, record, fields) => {
  row.classList.remove('due-today', 'overdue');
  if (closedStatuses.has(record.status)) return;

  const today = getLocalDateValue();
  const dates = fields
    .filter(field => field.type === 'date')
    .map(field => record[field.key])
    .filter(value => /^\d{4}-\d{2}-\d{2}$/.test(value || ''));

  if (dates.some(value => value === today)) row.classList.add('due-today');
  else if (dates.some(value => value < today)) row.classList.add('overdue');
};

const createRegister = ({ storageKey, rowsId, addId, fields, defaults }) => {
  const rowsContainer = document.querySelector(`#${rowsId}`);
  const addButton = document.querySelector(`#${addId}`);
  if (!rowsContainer || !addButton) return;

  const readRows = () => {
    try {
      const saved = JSON.parse(localStorage.getItem(storageKey));
      return Array.isArray(saved) && saved.length ? saved : cloneRows(defaults);
    } catch {
      return cloneRows(defaults);
    }
  };

  let rows = readRows();
  const saveRows = () => localStorage.setItem(storageKey, JSON.stringify(rows));

  const renderRows = () => {
    rowsContainer.replaceChildren();
    rows.forEach((record, index) => {
      const row = document.createElement('tr');
      fields.forEach(field => {
        const cell = document.createElement('td');
        let control;
        if (field.type === 'select') {
          control = document.createElement('select');
          (field.options || statusOptions).forEach(label => {
            const option = document.createElement('option');
            option.value = label;
            option.textContent = label;
            option.selected = record[field.key] === label;
            control.append(option);
          });
          control.addEventListener('change', () => {
            rows[index][field.key] = control.value;
            saveRows();
            updateRowDateState(row, rows[index], fields);
          });
        } else if (field.type === 'textarea') {
          control = document.createElement('textarea');
          control.rows = 2;
          control.value = record[field.key] ?? '';
          control.addEventListener('input', () => { rows[index][field.key] = control.value; saveRows(); });
        } else {
          control = document.createElement('input');
          control.type = field.type || 'text';
          control.value = record[field.key] ?? '';
          control.addEventListener('input', () => {
            rows[index][field.key] = control.value;
            saveRows();
            updateRowDateState(row, rows[index], fields);
          });
        }
        control.setAttribute('aria-label', `${field.label}, row ${index + 1}`);
        cell.append(control);
        row.append(cell);
      });
      updateRowDateState(row, record, fields);
      rowsContainer.append(row);
    });
  };

  addButton.addEventListener('click', () => {
    const emptyRow = {};
    fields.forEach(field => { emptyRow[field.key] = field.type === 'select' ? (field.options || statusOptions)[0] : ''; });
    rows.push(emptyRow);
    saveRows();
    renderRows();
    rowsContainer.lastElementChild?.querySelector('input, textarea, select')?.focus();
  });

  renderRows();
};

createRegister({
  storageKey: 'utar-yoga-exco-events-2026-2027-v1',
  rowsId: 'actionRows',
  addId: 'addAction',
  fields: [
    { key: 'event', label: 'Event or programme' },
    { key: 'period', label: 'Planned period' },
    { key: 'action', label: 'Decision or next action', type: 'textarea' },
    { key: 'lead', label: 'Lead' },
    { key: 'target', label: 'Target date', type: 'date' },
    { key: 'evidence', label: 'Approval or evidence', type: 'textarea' },
    { key: 'status', label: 'Status', type: 'select' }
  ],
  defaults: [
    { event: 'October Recruitment Drive', period: 'October 2026', action: 'Approve the recruitment offer, booth plan, vacancy message, merchandise display and follow-up survey.', lead: 'Higher Committee', target: '2026-09-25', evidence: 'DSA requirements; confirmed Treasurer and Secretary information', status: 'In progress' },
    { event: 'October free yoga class', period: 'October 2026', action: 'Confirm date, venue, instructor, first-time participant rule, registration method and attendance record.', lead: 'Programme Team', target: '2026-10-02', evidence: 'Approved proposal, venue and instructor confirmation', status: 'Proposed' },
    { event: 'Weekly yoga classes', period: 'Oct 2026–Feb 2027', action: 'Approve schedule, fees, attendance method and the semester or yearly certification model.', lead: 'Programme Team', target: '2026-09-30', evidence: 'Instructor availability, budget and certification policy', status: 'Proposed' },
    { event: 'Members’ Night and lucky draw', period: 'After October free class', action: 'Make a go/no-go decision and confirm eligibility, benefits, draw method, budget and event owner.', lead: 'EXCO', target: '2026-10-16', evidence: 'Budget approval and documented participation rules', status: 'Proposed' },
    { event: 'Student Instructor Programme', period: '2026/2027 programme year', action: 'Approve entry routes, attendance evidence, fee-waiver conditions and authorised certificate signatories.', lead: 'Student Coach HOD', target: '2026-10-15', evidence: 'Instructor, advisor and committee confirmation', status: 'Proposed' },
    { event: 'Yoga Club merchandise', period: 'October–November 2026', action: 'Complete the committee poll, shortlist products, obtain quotations and set a provisional budget ceiling.', lead: 'Publicity + Treasurer', target: '2026-10-30', evidence: 'Poll results, supplier quotations and funding decision', status: 'Not started' },
    { event: 'Meditation and yoga practice pilot', period: 'Semester 2026/2027', action: 'Confirm pilot scope, frequency, venue, fee, instructor eligibility, supervision and service-hour records.', lead: 'Programme Team', target: '2026-11-06', evidence: 'Advisor approval and documented pilot safeguards', status: 'Proposed' },
    { event: 'EXCO Meeting 2', period: 'Before Week 14', action: 'Schedule a programme-control review and close outstanding approvals, reports and action-register items.', lead: 'Chairperson + Secretary', target: '2026-11-20', evidence: 'Agenda, attendance list, minutes and updated registers', status: 'Not started' },
    { event: 'Natural Yoga pilot', period: 'February 2027 or later', action: 'Define the concept, facilitator, venue, demand test and resource requirements before approval.', lead: 'Programme Team', target: '2027-01-22', evidence: 'Feasibility note and advisor consultation', status: 'Deferred' },
    { event: 'Yoga Day', period: 'Future large-scale event', action: 'Consult Ms Belinda on institutional requirements and prepare an early feasibility brief.', lead: 'Chairperson', target: '2027-02-12', evidence: 'Consultation record, scope, capacity and preliminary budget', status: 'Deferred' }
  ]
});

createRegister({
  storageKey: 'utar-yoga-secretarial-register-2026-2027-v1',
  rowsId: 'secretarialRows',
  addId: 'addSecretarial',
  fields: [
    { key: 'document', label: 'Document' }, { key: 'category', label: 'Category' },
    { key: 'activity', label: 'Related activity' }, { key: 'requirement', label: 'Pre-defined requirement', type: 'textarea' },
    { key: 'owner', label: 'Owner' }, { key: 'target', label: 'Target' },
    { key: 'status', label: 'Status', type: 'select' }
  ],
  defaults: [
    { document: 'Recruitment Drive Proposal', category: 'Proposal', activity: 'October Recruitment Drive', requirement: 'Objectives, booth plan, offers, vacancy message, merchandise, budget, publicity and survey.', owner: 'Secretary', target: 'Before publicity approval', status: 'Drafting' },
    { document: 'Weekly Yoga Classes Proposal', category: 'Proposal', activity: 'Weekly classes', requirement: 'Schedule, venue, instructor, fee, registration, attendance and certification arrangements.', owner: 'Secretary', target: 'Before first class', status: 'Not started' },
    { document: 'Free Yoga Class Proposal', category: 'Proposal', activity: 'October free class', requirement: 'Date, venue, instructor, target participants, first-time offer, budget and attendance method.', owner: 'Secretary', target: 'Before promotion', status: 'Not started' },
    { document: 'Members’ Night and Lucky Draw Proposal', category: 'Proposal', activity: 'Members’ Night', requirement: 'Purpose, eligibility, benefits, draw method, budget, procurement, publicity and approval route.', owner: 'Secretary', target: 'After EXCO go decision', status: 'Proposed' },
    { document: 'Meditation / Practice Pilot Proposal', category: 'Proposal', activity: 'Pilot classes', requirement: 'Scope, frequency, venue, fee, student instructor eligibility, supervision and records.', owner: 'Secretary', target: 'Before recruitment', status: 'Proposed' },
    { document: 'Merchandise Decision Record', category: 'Approval record', activity: 'Yoga merchandise', requirement: 'Poll result, shortlist, design approval, supplier quotation, minimum order and funding decision.', owner: 'Secretary + Treasurer', target: 'Before order', status: 'Not started' },
    { document: 'EXCO Meeting 1 Minutes', category: 'Minutes', activity: 'EXCO Meeting 1', requirement: 'Attendance, declarations, decisions, action owners, due dates and approval dependencies.', owner: 'Secretary', target: 'Within 7 days', status: 'Drafting' },
    { document: 'EXCO Meeting 2 Minutes', category: 'Minutes', activity: 'EXCO Meeting 2', requirement: 'Progress review, outstanding decisions, closed actions and matters carried forward.', owner: 'Secretary', target: 'Within 7 days', status: 'Not started' },
    { document: 'Weekly Classes Semester Report', category: 'Report', activity: 'Weekly classes', requirement: 'Schedule delivered, attendance, finance, outcomes, certification and improvement points.', owner: 'Secretary', target: 'By Week 14', status: 'Not started' },
    { document: 'Recruitment Drive Report', category: 'Report', activity: 'Recruitment Drive', requirement: 'Participation, sign-ups, expenditure, publicity reach, survey findings and follow-up actions.', owner: 'Secretary', target: 'By Week 14', status: 'Not started' },
    { document: 'Free Yoga Class Report', category: 'Report', activity: 'Free yoga class', requirement: 'Attendance, participant profile, delivery, cost, feedback and conversion to membership/classes.', owner: 'Secretary', target: 'By Week 14', status: 'Not started' },
    { document: 'Office Bearers Report', category: 'Report', activity: 'Club governance', requirement: 'Confirmed roles, vacancies, appointments, service contributions and changes during the semester.', owner: 'Secretary', target: 'By Week 14', status: 'Not started' },
    { document: 'Master Document Checklist', category: 'Control record', activity: 'All activities', requirement: 'Document title, owner, due date, submission date, version, approval outcome and archive location.', owner: 'Secretary', target: 'Update weekly', status: 'In progress' }
  ]
});

createRegister({
  storageKey: 'utar-yoga-publicity-register-2026-2027-v1',
  rowsId: 'publicityRows',
  addId: 'addPublicity',
  fields: [
    { key: 'campaign', label: 'Campaign' }, { key: 'deliverable', label: 'Required deliverable', type: 'textarea' },
    { key: 'verification', label: 'Information to verify', type: 'textarea' }, { key: 'channel', label: 'Channel' },
    { key: 'window', label: 'Publish window' }, { key: 'owner', label: 'Owner' },
    { key: 'status', label: 'Status', type: 'select' }
  ],
  defaults: [
    { campaign: 'October Recruitment Drive', deliverable: 'Master poster, social caption, booth visual, QR registration and reminder post.', verification: 'Date, venue, booth, benefits, merchandise, lucky draw and approved registration link.', channel: 'Instagram, WhatsApp, campus channels', window: '2–3 weeks before event', owner: 'Publicity HOD', status: 'Drafting' },
    { campaign: 'Weekly yoga classes', deliverable: 'Semester schedule, class poster, recurring reminders and registration information.', verification: 'Instructor, dates, venue, fee, capacity, attendance and certificate wording.', channel: 'Instagram + WhatsApp', window: 'Before semester classes', owner: 'Publicity Team', status: 'Not started' },
    { campaign: 'October free yoga class', deliverable: 'Launch poster, registration post, reminder story and final-day notice.', verification: 'First-time offer, eligibility, instructor, venue, capacity and closing time.', channel: 'Instagram + recruitment channels', window: '7–14 days before class', owner: 'Publicity Team', status: 'Not started' },
    { campaign: 'Members’ Night and lucky draw', deliverable: 'Event announcement, eligibility explainer, benefit visual and winner communication.', verification: 'Go decision, eligible audience, budgeted benefits, entry method and draw governance.', channel: 'Member channels', window: 'Only after approval', owner: 'Publicity HOD', status: 'Proposed' },
    { campaign: 'Vice Treasurer vacancy', deliverable: 'Role card, expression-of-interest message and application route.', verification: 'Role scope, eligibility, selection owner, deadline and contact person.', channel: 'Recruitment booth + social', window: 'Recruitment Drive cycle', owner: 'Publicity + Treasurer', status: 'In progress' },
    { campaign: 'Student Instructor Programme', deliverable: 'Programme explainer, eligibility guide and expression-of-interest form.', verification: 'Entry route, attendance threshold, fee benefit, service expectations and signatories.', channel: 'Member channels', window: 'After policy approval', owner: 'Publicity + Student Coach HOD', status: 'Proposed' },
    { campaign: 'Yoga merchandise', deliverable: 'Committee poll, approved concept previews, price notice and order instructions.', verification: 'Shortlist, design, unit price, minimum order, payment deadline and fulfilment plan.', channel: 'Internal poll then social', window: 'October–November 2026', owner: 'Publicity Team', status: 'Not started' },
    { campaign: 'Meditation / practice pilot', deliverable: 'Pilot invitation, session information, registration link and feedback request.', verification: 'Approved scope, facilitator, supervision, venue, fee, capacity and service evidence.', channel: 'Member channels', window: 'After pilot approval', owner: 'Publicity Team', status: 'Proposed' },
    { campaign: 'Post-event follow-up survey', deliverable: 'Survey link, thank-you message, response reminder and findings summary.', verification: 'Survey owner, questions, privacy statement, closing date and reporting audience.', channel: 'Participant channels', window: 'Within 24 hours after event', owner: 'Publicity + Secretary', status: 'Not started' },
    { campaign: 'Publicity archive and corrections', deliverable: 'Final asset archive and correction log for superseded information.', verification: 'Final approval source, publication date, file name, version and affected channels.', channel: 'Club drive / website', window: 'After every publication', owner: 'Publicity HOD', status: 'In progress' }
  ]
});

createRegister({
  storageKey: 'utar-yoga-future-planning-2026-2027-v1',
  rowsId: 'futureRows',
  addId: 'addFuture',
  fields: [
    { key: 'opportunity', label: 'Future opportunity' }, { key: 'question', label: 'Question to resolve', type: 'textarea' },
    { key: 'evidence', label: 'Evidence or consultation', type: 'textarea' }, { key: 'owner', label: 'Exploration owner' },
    { key: 'review', label: 'Review date', type: 'date' }, { key: 'status', label: 'Stage', type: 'select', options: ['Idea', 'Exploring', 'Ready for decision', 'Approved', 'Deferred'] }
  ],
  defaults: [
    { opportunity: 'Yoga Day', question: 'What scale, institutional approvals, capacity and funding would a larger event require?', evidence: 'Consult Ms Belinda; prepare scope and preliminary budget.', owner: 'Chairperson', review: '2027-02-12', status: 'Exploring' },
    { opportunity: 'Natural Yoga', question: 'What does the programme include, who facilitates it and when is demand strongest?', evidence: 'Concept note, facilitator options and member interest check.', owner: 'Programme Team', review: '2027-01-22', status: 'Idea' },
    { opportunity: 'Club website as authoritative record', question: 'Which financial and participation records should move into the website, and who controls access?', evidence: 'Process map, access review and Treasurer/Secretary consultation.', owner: 'Treasurer + Secretary', review: '2027-01-29', status: 'Exploring' },
    { opportunity: 'Student instructor pipeline', question: 'How will the club identify, train, supervise and retain future student instructors?', evidence: 'Competency criteria, mentoring capacity and certification pathway.', owner: 'Student Coach HOD', review: '2027-02-05', status: 'Idea' },
    { opportunity: 'External partnerships', question: 'Which partners could strengthen facilities, expertise or member benefits without compromising club governance?', evidence: 'Partner shortlist, value exchange, risks and approval requirements.', owner: 'Higher Committee', review: '2027-02-19', status: 'Idea' }
  ]
});
