import { useState, useEffect, useRef, createContext, useContext, Component } from "react";
import {
  Users, Calendar, Wallet, LogOut, Plus, Trash2, CheckCircle2,
  XCircle, Eye, EyeOff, UserPlus, ShieldCheck, ClipboardList, TrendingDown,
  MoreVertical, Copy, Check, CheckCheck, KeyRound, Settings, Lock, X, Palette, Type,
  Camera, Globe, User as UserIcon, ChevronDown, Sun, Moon, ChevronLeft, ChevronRight,
  Menu, ChevronUp, UserX, ArrowLeft, Paintbrush, Download, Send, Bell, Search, LayoutDashboard, Home, Share2
} from "lucide-react";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";
import confetti from "canvas-confetti";

// FIX: avvalgi versiya new Date().toISOString() orqali UTC sanasini olardi.
// O'zbekiston UTC+5 da, shu sabab tungi soat 00:00–05:00 orasida (mahalliy vaqt)
// "bugungi sana" noto'g'ri — aslida kechagi kun qaytardi. Endi mahalliy vaqt asosida hisoblaymiz.
const todayISO = () => {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};
const fmt = (n) => Number(n || 0).toLocaleString("uz-UZ") + " so'm";
const fmtDays = (n) => {
  const r = Math.round(Number(n || 0) * 10) / 10;
  return Number.isInteger(r) ? String(r) : r.toFixed(1);
};

// YANGI: asosiy tugmalar uchun accent rangidan yumshoq gradient hosil qiladi —
// login ekranidagi gradient uslubi bilan uyg'unlik uchun.
function darkenHex(hex, amount) {
  const h = hex.replace("#", "");
  const num = parseInt(h, 16);
  const delta = Math.round(255 * amount);
  const r = Math.max(0, (num >> 16) - delta);
  const g = Math.max(0, ((num >> 8) & 0x00ff) - delta);
  const b = Math.max(0, (num & 0x0000ff) - delta);
  return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0")}`;
}
function accentGradient(hex) {
  return `linear-gradient(155deg, ${hex} 0%, ${darkenHex(hex, 0.28)} 100%)`;
}

function wageForDate(emp, date) {
  const history = emp.wageHistory;
  if (!Array.isArray(history) || history.length === 0) return Number(emp.dailyWage || 0);
  let applicable = history[0];
  for (const entry of history) {
    if (entry.date <= date) applicable = entry;
    else break;
  }
  return Number(applicable.wage);
}

function attEntryStatus(raw) {
  if (raw && typeof raw === "object") return raw.v;
  if (typeof raw === "number") return raw;
  if (raw === true) return 1;
  return 0;
}
function attEntryWage(raw, emp, date) {
  if (raw && typeof raw === "object" && typeof raw.wage === "number") return raw.wage;
  return wageForDate(emp, date);
}

// YANGI: bazadagi "settlements" qatorini ilova ichida ishlatiladigan ko'rinishga o'tkazadi.
function mapSettlementRow(r) {
  return {
    id: r.id,
    closedThrough: r.closed_through,
    openingBalance: Number(r.opening_balance || 0),
    workedDays: Number(r.worked_days || 0),
    totalWage: Number(r.total_wage || 0),
    totalPaid: Number(r.total_paid || 0),
    remaining: Number(r.remaining || 0),
    carryOver: Number(r.carry_over || 0),
  };
}

function employeeJoinDate(emp) {
  if (Array.isArray(emp.wageHistory) && emp.wageHistory.length > 0) return emp.wageHistory[0].date;
  return "2000-01-01"; // eski (tarixsiz) ishchilar uchun — hamma joyda ko'rinaveradi
}

const LANGS = [
  { code: "uz", label: "O'zbekcha" },
  { code: "ru", label: "Русский" },
  { code: "en", label: "English" },
];

const STR = {
  appName: { uz: "Nazorat+", ru: "Nazorat+", en: "Nazorat+" },
  loginTitle: { uz: "Tizimga kirish", ru: "Вход в систему", en: "Sign in" },
  loginHeading: { uz: "Kirish", ru: "Вход", en: "Login" },
  loginSubtitle: { uz: "Login va parolingizni kiriting", ru: "Введите логин и пароль", en: "Enter your login and password" },
  login: { uz: "Login", ru: "Логин", en: "Username" },
  password: { uz: "Parol", ru: "Пароль", en: "Password" },
  loginBtn: { uz: "Kirish", ru: "Войти", en: "Sign in" },
  firstTimeHint: { uz: "Birinchi marta kiryapsizmi? Admin uchun:", ru: "Первый раз здесь? Для админа:", en: "First time here? For admin:" },
  changeLaterHint: { uz: "Keyinchalik parolni albatta o'zgartiring.", ru: "Обязательно смените пароль позже.", en: "Be sure to change the password later." },
  wrongLogin: { uz: "Login yoki parol noto'g'ri", ru: "Неверный логин или пароль", en: "Incorrect username or password" },
  logout: { uz: "Chiqish", ru: "Выход", en: "Log out" },
  adminPanel: { uz: "Tizim", ru: "Система", en: "System" },
  employeePanel: { uz: "Profil", ru: "Профиль", en: "Profile" },
  navEmployees: { uz: "Ishchilar", ru: "Сотрудники", en: "Employees" },
  navAttendance: { uz: "Davomat", ru: "Посещаемость", en: "Attendance" },
  navAdvances: { uz: "Avanslar", ru: "Авансы", en: "Advances" },
  navReport: { uz: "Hisobot", ru: "Отчёт", en: "Report" },
  addEmployeeHeader: { uz: "Yangi ishchi qo'shish", ru: "Добавить сотрудника", en: "Add new employee" },
  fullName: { uz: "Ism familiya", ru: "Имя фамилия", en: "Full name" },
  dailyWage: { uz: "Kunlik ish haqi (so'm)", ru: "Дневная зарплата (сум)", en: "Daily wage" },
  add: { uz: "Qo'shish", ru: "Добавить", en: "Add" },
  noEmployees: { uz: "Hali ishchilar qo'shilmagan", ru: "Сотрудники ещё не добавлены", en: "No employees added yet" },
  perDay: { uz: "/kun", ru: "/день", en: "/day" },
  daysWorkedSuffix: { uz: "kun ishlagan", ru: "дней отработано", en: "days worked" },
  details: { uz: "Batafsil", ru: "Подробнее", en: "Details" },
  credentialsHeader: { uz: "Kirish ma'lumotlari", ru: "Данные для входа", en: "Login credentials" },
  copy: { uz: "Nusxalash", ru: "Копировать", en: "Copy" },
  copied: { uz: "Nusxalandi", ru: "Скопировано", en: "Copied" },
  day: { uz: "Kun", ru: "Дни", en: "Days" },
  advance: { uz: "To'langan", ru: "Выплачено", en: "Paid" },
  remaining: { uz: "Qoldiq", ru: "Остаток", en: "Remaining" },
  deleteEmployee: { uz: "Ishchini o'chirish", ru: "Удалить сотрудника", en: "Delete employee" },
  confirmDelete: { uz: "Rostdan ham {name}ni o'chirmoqchimisiz? Davomat va avans tarixi ham butunlay o'chib ketadi.", ru: "Точно удалить {name}? История посещаемости и авансов тоже удалится.", en: "Really delete {name}? Attendance and advance history will be deleted too." },
  yesDelete: { uz: "Ha, o'chirish", ru: "Да, удалить", en: "Yes, delete" },
  yesConfirm: { uz: "Ha, qo'llash", ru: "Да, применить", en: "Yes, apply" },
  cancel: { uz: "Bekor qilish", ru: "Отмена", en: "Cancel" },
  markAttendanceHeader: { uz: "Davomatni belgilash", ru: "Отметить посещаемость", en: "Mark attendance" },
  selectPlaceholder: { uz: "Tanlang...", ru: "Выберите...", en: "Select..." },
  employee: { uz: "Ishchi", ru: "Сотрудник", en: "Employee" },
  date: { uz: "Sana", ru: "Дата", en: "Date" },
  present: { uz: "Ishga keldi", ru: "Пришёл", en: "Present" },
  fullDay: { uz: "To'liq kun", ru: "Полный день", en: "Full day" },
  halfDay: { uz: "Yarim kun", ru: "Полдня", en: "Half day" },
  absent: { uz: "Kelmadi", ru: "Не пришёл", en: "Absent" },
  clear: { uz: "Tozalash", ru: "Очистить", en: "Clear" },
  statusPrefix: { uz: "holati:", ru: "статус:", en: "status:" },
  statusFull: { uz: "to'liq kun ishlagan", ru: "отработал полный день", en: "worked full day" },
  statusHalf: { uz: "yarim kun ishlagan", ru: "отработал полдня", en: "worked half day" },
  statusAbsent: { uz: "kelmagan", ru: "отсутствовал", en: "absent" },
  statusNone: { uz: "belgilanmagan", ru: "не отмечено", en: "not marked" },
  totalWorkedDays: { uz: "Jami ishlagan kunlar:", ru: "Всего отработано дней:", en: "Total days worked:" },
  giveAdvanceHeader: { uz: "Avans / to'lov berish", ru: "Выдать аванс / оплату", en: "Give an advance / payment" },
  typeAvans: { uz: "Avans", ru: "Аванс", en: "Advance" },
  typeSalary: { uz: "Ish haqi to'lovi", ru: "Выплата зарплаты", en: "Salary payment" },
  addSalaryPayment: { uz: "To'lovni qo'shish", ru: "Добавить выплату", en: "Add payment" },
  amount: { uz: "Summa (so'm)", ru: "Сумма (сум)", en: "Amount" },
  note: { uz: "Izoh (ixtiyoriy)", ru: "Заметка (необязательно)", en: "Note (optional)" },
  addAdvance: { uz: "Avans qo'shish", ru: "Добавить аванс", en: "Add advance" },
  advanceHistory: { uz: "Avanslar tarixi", ru: "История авансов", en: "Advance history" },
  noAdvances: { uz: "Hali avans olinmagan", ru: "Авансов ещё не было", en: "No advances yet" },
  reportHeader: { uz: "Barcha ishchilar bo'yicha hisobot", ru: "Отчёт по всем сотрудникам", en: "Report for all employees" },
  colEmployee: { uz: "Ishchi", ru: "Сотрудник", en: "Employee" },
  colDays: { uz: "Kun", ru: "Дни", en: "Days" },
  colCalculated: { uz: "Hisoblangan", ru: "Начислено", en: "Calculated" },
  colAdvance: { uz: "To'langan", ru: "Выплачено", en: "Paid" },
  colRemaining: { uz: "Qoldiq", ru: "Остаток", en: "Remaining" },
  noData: { uz: "Ma'lumot yo'q", ru: "Нет данных", en: "No data" },
  myWorkedDays: { uz: "Ishlagan kunlarim", ru: "Мои отработанные дни", en: "My worked days" },
  noAttendanceYet: { uz: "Hali davomat belgilanmagan", ru: "Посещаемость ещё не отмечена", en: "No attendance marked yet" },
  futureDateWarning: { uz: "Hali kelmagan sanaga davomat belgilab bo'lmaydi", ru: "Нельзя отметить посещаемость на будущую дату", en: "You can't mark attendance for a future date" },
  markAllFull: { uz: "Hammasi to'liq", ru: "Все полный день", en: "Mark all full" },
  markAllHalf: { uz: "Hammasi yarim", ru: "Все полдня", en: "Mark all half" },
  markAllAbsent: { uz: "Hammasi kelmadi", ru: "Все отсутствуют", en: "Mark all absent" },
  myAdvances: { uz: "Avanslarim", ru: "Мои авансы", en: "My advances" },
  noAdvancesYet: { uz: "Hali avans olmagansiz", ru: "Вы ещё не получали аванс", en: "You haven't taken an advance yet" },
  statWorkedDays: { uz: "Ishlagan kunlar", ru: "Отработано дней", en: "Days worked" },
  statDailyWage: { uz: "Kunlik stavka", ru: "Дневная ставка", en: "Daily rate" },
  statTakenAdvance: { uz: "Olingan avans", ru: "Взятый аванс", en: "Advance taken" },
  statRemainingSalary: { uz: "Qolgan maosh", ru: "Остаток зарплаты", en: "Remaining salary" },
  profile: { uz: "Profil", ru: "Профиль", en: "Profile" },
  changePhoto: { uz: "Rasm qo'yish", ru: "Загрузить фото", en: "Upload photo" },
  security: { uz: "Xavfsizlik", ru: "Безопасность", en: "Security" },
  currentPassword: { uz: "Joriy parol", ru: "Текущий пароль", en: "Current password" },
  newLogin: { uz: "Yangi login", ru: "Новый логин", en: "New username" },
  newPassword: { uz: "Yangi parol", ru: "Новый пароль", en: "New password" },
  repeatNewPassword: { uz: "Yangi parolni takrorlang", ru: "Повторите новый пароль", en: "Repeat new password" },
  save: { uz: "Saqlash", ru: "Сохранить", en: "Save" },
  editWage: { uz: "Ish haqini o'zgartirish", ru: "Изменить зарплату", en: "Edit wage" },
  newDailyWage: { uz: "Yangi kunlik ish haqi (so'm)", ru: "Новая дневная зарплата (сум)", en: "New daily wage" },
  language: { uz: "Til", ru: "Язык", en: "Language" },
  themeColor: { uz: "Mavzu rangi", ru: "Цвет темы", en: "Theme color" },
  lightBg: { uz: "Oq fon", ru: "Светлый фон", en: "Light background" },
  darkBg: { uz: "Qora fon", ru: "Тёмный фон", en: "Dark background" },
  fontSize: { uz: "Shrift o'lchami", ru: "Размер шрифта", en: "Font size" },
  fontSmall: { uz: "Kichik", ru: "Мелкий", en: "Small" },
  fontMedium: { uz: "O'rta", ru: "Средний", en: "Medium" },
  fontLarge: { uz: "Katta", ru: "Крупный", en: "Large" },
  errWrongCurrentPassword: { uz: "Joriy parol noto'g'ri", ru: "Текущий пароль неверен", en: "Current password is incorrect" },
  errEmptyLogin: { uz: "Login bo'sh bo'lmasin", ru: "Логин не может быть пустым", en: "Username can't be empty" },
  errShortPassword: { uz: "Yangi parol kamida 4 belgidan iborat bo'lsin", ru: "Новый пароль должен быть не короче 4 символов", en: "New password must be at least 4 characters" },
  errPasswordMismatch: { uz: "Yangi parollar bir xil emas", ru: "Новые пароли не совпадают", en: "New passwords don't match" },
  errLoginTaken: { uz: "Bu login band, boshqasini tanlang", ru: "Этот логин занят, выберите другой", en: "This username is taken, choose another" },
  savedOk: { uz: "Saqlandi! Endi yangi login-parol bilan kiring.", ru: "Сохранено! Теперь входите с новым логином и паролем.", en: "Saved! Sign in with the new username and password now." },
  fillAllFields: { uz: "Barcha maydonlarni to'ldiring", ru: "Заполните все поля", en: "Fill in all fields" },
  loading: { uz: "Yuklanmoqda...", ru: "Загрузка...", en: "Loading..." },
  admin: { uz: "Admin", ru: "Админ", en: "Admin" },
  enterAsAdmin: { uz: "Boshqaruvchi sifatida kirish", ru: "Войти как руководитель", en: "Sign in as manager" },
  enterAsEmployee: { uz: "Ishchi sifatida kirish", ru: "Войти как сотрудник", en: "Sign in as employee" },
  backToEmployeeLogin: { uz: "Ishchi kirishiga qaytish", ru: "Назад ко входу сотрудника", en: "Back to employee sign in" },
  adminLoginTitle: { uz: "Boshqaruvchi kirishi", ru: "Вход руководителя", en: "Manager sign in" },
  employeeLoginTitle: { uz: "Ishchi kirishi", ru: "Вход сотрудника", en: "Employee sign in" },
  registerTitle: { uz: "Yangi boshqaruvchi yaratish", ru: "Создать нового руководителя", en: "Create a new manager account" },
  roleTabEmployee: { uz: "Ishchi", ru: "Сотрудник", en: "Employee" },
  roleTabAdmin: { uz: "Boshqaruvchi", ru: "Руководитель", en: "Manager" },
  registerSubtitle: { uz: "O'z login-parolingizni o'ylab toping va o'z ishchilaringizni boshqaring", ru: "Придумайте свой логин и пароль и управляйте своими сотрудниками", en: "Choose your own username and password to manage your own employees" },
  chooseLogin: { uz: "Login o'ylab toping", ru: "Придумайте логин", en: "Choose a username" },
  choosePassword: { uz: "Parol o'ylab toping", ru: "Придумайте пароль", en: "Choose a password" },
  createAccountBtn: { uz: "Ro'yxatdan o'tish", ru: "Зарегистрироваться", en: "Register" },
  alreadyHaveAccount: { uz: "Mavjud hisobga kirish", ru: "Войти в существующий аккаунт", en: "Sign in to an existing account" },
  newHere: { uz: "Birinchi marta kiryapsizmi?", ru: "Впервые здесь?", en: "First time here?" },
  deleteAccount: { uz: "Akkauntni o'chirish", ru: "Удалить аккаунт", en: "Delete account" },
  confirmDeleteAccountAdmin: { uz: "Rostdan ham akkauntingizni o'chirmoqchimisiz? Barcha ishchilaringiz, davomat va avans tarixi ham butunlay o'chib ketadi.", ru: "Точно удалить аккаунт? Все ваши сотрудники, посещаемость и авансы тоже будут удалены навсегда.", en: "Really delete your account? All your employees, attendance, and advance history will be permanently deleted too." },
  confirmDeleteAccountEmployee: { uz: "Rostdan ham akkauntingizni o'chirmoqchimisiz? Davomat va avans tarixingiz ham butunlay o'chib ketadi.", ru: "Точно удалить аккаунт? Ваша посещаемость и авансы тоже будут удалены навсегда.", en: "Really delete your account? Your attendance and advance history will be permanently deleted too." },
  yesDeleteAccount: { uz: "Ha, akkauntni o'chirish", ru: "Да, удалить аккаунт", en: "Yes, delete account" },
  appearance: { uz: "Ko'rinish", ru: "Внешний вид", en: "Appearance" },
  privacySecurity: { uz: "Maxfiylik va xavfsizlik", ru: "Конфиденциальность и безопасность", en: "Privacy and Security" },
  updateCredentials: { uz: "Login va parolni yangilash", ru: "Обновить логин и пароль", en: "Update login and password" },
  enableNotifications: { uz: "Bildirishnomalarni yoqish", ru: "Включить уведомления", en: "Enable notifications" },
  notifications: { uz: "Bildirishnomalar", ru: "Уведомления", en: "Notifications" },
  noNotifications: { uz: "Hali bildirishnoma yo'q", ru: "Уведомлений пока нет", en: "No notifications yet" },
  markAllRead: { uz: "Hammasini o'qilgan deb belgilash", ru: "Отметить все как прочитанные", en: "Mark all as read" },
  advanced: { uz: "Akkaunt boshqaruvi", ru: "Управление аккаунтом", en: "Account management" },
  rememberMe: { uz: "Meni eslab qol", ru: "Запомнить меня", en: "Remember me" },
  forgotPassword: { uz: "Parolni unutdingizmi?", ru: "Забыли пароль?", en: "Forgot password?" },
  forgotPasswordHint: { uz: "Parolni faqat boshqaruvchi tiklashi mumkin. Iltimos, boshqaruvchingizga murojaat qiling.", ru: "Пароль может восстановить только руководитель. Пожалуйста, обратитесь к нему.", en: "Only your manager can reset your password. Please contact them." },
  noAccountYet: { uz: "Hisobingiz yo'qmi?", ru: "Нет аккаунта?", en: "Don't have an account?" },
  signUpLink: { uz: "Ro'yxatdan o'tish", ru: "Регистрация", en: "Sign up" },
};

function makeT(lang) {
  return (key, vars) => {
    let s = (STR[key] && (STR[key][lang] || STR[key].uz)) || key;
    if (vars) Object.entries(vars).forEach(([k, v]) => { s = s.replace(`{${k}}`, v); });
    return s;
  };
}

const AppContext = createContext({ accent: "#4d84d9", lang: "uz", t: makeT("uz") });
const useApp = () => useContext(AppContext);

const ACCENT_PRESETS = [
  { name: "Bronza", value: "#c98a4b" },
  { name: "Feruza", value: "#3a9188" },
  { name: "Ko'k", value: "#4d84d9" },
  { name: "Binafsha", value: "#8f7bd6" },
  { name: "Lolaqizg'aldoq", value: "#d9635a" },
  { name: "Pushti", value: "#cf6f98" },
];

const FONT_SCALES = [
  { key: "fontSmall", value: 90 },
  { key: "fontMedium", value: 100 },
  { key: "fontLarge", value: 115 },
];

const PALETTES = {
  dark: {
    "--bg-app": "#12151a",
    "--bg-card": "#1a1e25",
    "--bg-panel": "#181c22",
    "--border": "#262b33",
    "--border-input": "#333944",
    "--border-soft": "#20242b",
    "--text-primary": "#edeff2",
    "--text-secondary": "#9aa3af",
    "--text-muted": "#6b7280",
    "--text-faint": "#4d5560",
    "--shadow-card": "0 1px 2px rgba(0,0,0,0.35), 0 4px 10px rgba(0,0,0,0.25)",
    "--shadow-pop": "0 12px 32px rgba(0,0,0,0.45)",
  },
  light: {
    "--bg-app": "#f4f5f7",
    "--bg-card": "#ffffff",
    "--bg-panel": "#ffffff",
    "--border": "#e7e9ed",
    "--border-input": "#d7dbe1",
    "--border-soft": "#edeef1",
    "--text-primary": "#1d2229",
    "--text-secondary": "#5c6472",
    "--text-muted": "#889099",
    "--text-faint": "#a7adb6",
    "--shadow-card": "0 1px 2px rgba(16,24,40,0.04), 0 2px 8px rgba(16,24,40,0.06)",
    "--shadow-pop": "0 16px 40px rgba(16,24,40,0.14)",
  },
};

const VAPID_PUBLIC_KEY = "BJHw7YqggwDUKfjS9pOcZyA_y7MO_46FWaRKv-fF8zr71CDEycK7hlEzq_hq4IzW7VhzysMJFZ-jXP4ULEYzn3k";

function urlBase64ToUint8Array(base64String) {
  const padding = "=".repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
  const rawData = atob(base64);
  const outputArray = new Uint8Array(rawData.length);
  for (let i = 0; i < rawData.length; i++) outputArray[i] = rawData.charCodeAt(i);
  return outputArray;
}

// FIX: yangi qo'shildi — avval Error Boundary umuman yo'q edi, shu sabab har
// qanday kutilmagan JS xatosi (masalan null obyektdan xususiyat o'qishga urinish)
// butun ilovani oq ekranga aylantirib qo'yardi. Endi shunday xato yuz bersa,
// foydalanuvchi tushunarli xabar va "Qayta yuklash" tugmasini ko'radi.
class AppErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() {
    return { hasError: true };
  }
  componentDidCatch(error, info) {
    console.error("Ilovada kutilmagan xato:", error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16, background: "#101317", color: "#edeff2", padding: 24, textAlign: "center" }}>
          <div style={{ fontSize: 16, fontWeight: 600 }}>Kutilmagan xato yuz berdi</div>
          <div style={{ fontSize: 13, color: "#8d97a3", maxWidth: 320 }}>
            Ilova ishlashda muammo yuz berdi. Sahifani qayta yuklab ko'ring.
          </div>
          <button
            type="button"
            onClick={() => window.location.reload()}
            style={{ padding: "10px 20px", borderRadius: 8, background: "#4d84d9", color: "#12161c", fontWeight: 600, fontSize: 14, border: "none", cursor: "pointer" }}
          >
            Qayta yuklash
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

// YANGI: supabase.functions.invoke() xato qaytarganda, Supabase JS kutubxonasi
// odatda umumiy "Edge Function returned a non-2xx status code" kabi tushunarsiz
// xabar beradi va funksiyaning o'zi yozgan aniq xabarni (masalan "Bu login
// band") yashirib qo'yadi. Bu yordamchi funksiya asl JSON javobni o'qib,
// foydalanuvchiga tushunarli xabarni chiqarib beradi.
async function invokeFn(name, body) {
  const { data, error } = await supabase.functions.invoke(name, body !== undefined ? { body } : undefined);
  if (error) {
    let message = error.message || String(error);
    try {
      if (error.context && typeof error.context.json === "function") {
        const body2 = await error.context.json();
        if (body2 && body2.error) message = body2.error;
      }
    } catch (e) {}
    return { data: null, error: message };
  }
  if (data && data.error) return { data: null, error: data.error };
  return { data, error: null };
}

function fileToAvatarDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read failed"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("image failed"));
      img.onload = () => {
        const size = 200;
        const canvas = document.createElement("canvas");
        canvas.width = size; canvas.height = size;
        const ctx = canvas.getContext("2d");
        const minSide = Math.min(img.width, img.height);
        const sx = (img.width - minSide) / 2;
        const sy = (img.height - minSide) / 2;
        ctx.drawImage(img, sx, sy, minSide, minSide, 0, 0, size, size);
        resolve(canvas.toDataURL("image/jpeg", 0.82));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function Field({ label, value, onChange, type = "text" }) {
  const [reveal, setReveal] = useState(false);
  const isPassword = type === "password";
  return (
    <div>
      <label className="block text-xs text-[var(--text-secondary)] mb-1.5">{label}</label>
      <div className={isPassword ? "relative" : ""}>
        <input
          type={isPassword ? (reveal ? "text" : "password") : type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          className={`w-full px-3 py-2.5 ${isPassword ? "pr-10" : ""} rounded-lg field text-[var(--text-primary)] text-sm outline-none focus:border-[var(--accent)] transition-colors`}
        />
        {isPassword && (
          <button
            type="button"
            onClick={() => setReveal((v) => !v)}
            className="absolute right-0 top-0 h-full px-3 flex items-center text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors"
          >
            {reveal ? <EyeOff size={15} /> : <Eye size={15} />}
          </button>
        )}
      </div>
    </div>
  );
}

// YANGI: rasmga o'xshab, dumaloq (pill) ko'rinishdagi, ichida ikonka bo'lgan input.
// Login ekranida foydalaniladi — label yo'q, o'rniga placeholder ishlatiladi.
function IconInput({ icon, type = "text", value, onChange, placeholder, onKeyDown, autoFocus, showToggle, toggleIcon, onToggle, error, disabled, shakeKey }) {
  const [focused, setFocused] = useState(false);
  const baseShadow = "inset -6px -6px 10px rgba(255,255,255,0.95), inset 6px 6px 10px rgba(184,190,204,0.45)";
  const focusShadow = "inset 3px 3px 6px rgba(20,60,140,0.35), inset -3px -3px 6px rgba(70,130,220,0.25)";
  const errorShadow = "inset 3px 3px 6px rgba(209,69,59,0.4), inset -3px -3px 6px rgba(220,100,90,0.3)";
  return (
    <div key={shakeKey} className={`relative ${error ? "shake-anim" : ""}`}>
      <span
        className="absolute left-4 top-1/2 -translate-y-1/2 pointer-events-none transition-colors duration-300"
        style={{ color: error ? "#a10f0f" : focused ? "#1a56b0" : "#909090" }}
      >
        {icon}
      </span>
      <input
        type={type}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={onKeyDown}
        placeholder={placeholder}
        autoFocus={autoFocus}
        disabled={disabled}
        onFocus={() => setFocused(true)}
        onBlur={() => setFocused(false)}
        className={`w-full pl-11 ${showToggle ? "pr-11" : "pr-4"} py-3.5 rounded-2xl bg-[#e8e8e8] text-[#4a4a4a] text-sm font-medium outline-none border-none transition-all duration-300 placeholder:text-[#a3a3a3] disabled:opacity-60`}
        style={{ boxShadow: error ? errorShadow : focused ? focusShadow : baseShadow, transform: focused ? "translateY(-2px)" : "translateY(0)" }}
      />
      {showToggle && (
        <button
          type="button"
          onClick={onToggle}
          className="absolute right-4 top-1/2 -translate-y-1/2 text-[#909090] hover:text-[#a10f0f] transition-colors"
        >
          {toggleIcon}
        </button>
      )}
    </div>
  );
}
// YANGI: "Meni eslab qol" uchun kichik dumaloq svitch (toggle).
function ToggleSwitch({ checked, onChange }) {
  return (
    <button
      type="button"
      onClick={() => onChange(!checked)}
      className="relative w-8 h-[18px] rounded-full shrink-0 bg-[#e8e8e8] transition-all duration-300"
      style={{
        boxShadow: checked
          ? "inset 3px 3px 6px rgba(20,60,140,0.35), inset -3px -3px 6px rgba(70,130,220,0.25)"
          : "inset 3px 3px 6px rgba(184,190,204,0.5), inset -3px -3px 6px rgba(255,255,255,0.9)",
      }}
      aria-pressed={checked}
    >
      <span
        className="absolute top-0.5 w-3.5 h-3.5 rounded-full transition-all duration-300"
        style={{
          left: checked ? "16px" : "2px",
          backgroundColor: checked ? "#1a56b0" : "#e8e8e8",
          boxShadow: "1px 1px 3px rgba(0,0,0,0.25)",
        }}
      />
    </button>
  );
}

function MoneyField({ label, value, onChange, suffix }) {
  const digits = String(value || "").replace(/\D/g, "");
  const display = digits ? Number(digits).toLocaleString("uz-UZ") : "";
  return (
    <div>
      <label className="block text-xs text-[var(--text-secondary)] mb-1.5">{label}</label>
      <div className="relative">
        <input
          type="text"
          inputMode="numeric"
          value={display}
          onChange={(e) => onChange(e.target.value.replace(/\D/g, ""))}
          placeholder="0"
          className={`w-full px-3 py-2.5 ${suffix ? "pr-14" : ""} rounded-lg field text-[var(--text-primary)] text-sm font-mono tabular-nums outline-none focus:border-[var(--accent)] transition-colors`}
        />
        {suffix && (
          <span className="absolute right-3 top-0 h-full flex items-center text-[var(--text-muted)] text-xs">{suffix}</span>
        )}
      </div>
    </div>
  );
}

// YANGI: bir nechta statistikani alohida-alohida kartochkalarga o'rash o'rniga,
// bitta umumiy kartochka ichida chiziqlar bilan ajratilgan katakchalarga
// joylashtiradi — ekranda "kartochkalar to'lib ketishi"ning oldini oladi.
function StatCell({ label, value, tone = "default", icon, border = "" }) {
  const toneMap = {
    default: "text-[var(--text-primary)]",
    good: "text-[var(--good)]",
    bad: "text-[var(--bad)]",
  };
  const borderClass = [
    border.includes("r") ? "border-r" : "",
    border.includes("b") ? "border-b" : "",
  ].filter(Boolean).join(" ");
  return (
    <div className={`p-4 min-w-0 overflow-hidden ${borderClass} border-[var(--border-soft)]`}>
      <div className="flex items-center gap-1.5 text-[var(--text-muted)] text-[11px] mb-1.5 leading-snug">
        {icon}<span>{label}</span>
      </div>
      <div className={`text-base font-semibold font-mono tabular-nums leading-tight whitespace-nowrap truncate ${toneMap[tone]}`}>{value}</div>
    </div>
  );
}

// YANGI: bo'sh holatlar uchun — shunchaki matn o'rniga ikonka + (kerak bo'lsa)
// harakat tugmasi bilan, ekran "bo'sh" his qilinmasligi uchun.
function EmptyState({ icon, text, actionLabel, onAction }) {
  const { accent } = useApp();
  return (
    <div className="flex flex-col items-center justify-center gap-2.5 py-10 text-center">
      <div className="w-11 h-11 rounded-full flex items-center justify-center bg-[var(--bg-app)] text-[var(--text-faint)]">
        {icon}
      </div>
      <p className="text-[var(--text-muted)] text-sm">{text}</p>
      {actionLabel && onAction && (
        <button
          type="button"
          onClick={onAction}
          className="mt-1 flex items-center gap-1.5 px-4 py-2 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity"
          style={{ background: accentGradient(accent) }}
        >
          <Plus size={14} /> {actionLabel}
        </button>
      )}
    </div>
  );
}

// YANGI: pul summasi o'zgarganda eski qiymatdan yangi qiymatgacha animatsiyali
// sanaydi — faqat 1-2 ta muhim joyda (umumiy qoldiq, jami qarz) ishlatiladi.
function AnimatedAmount({ value, formatter, className }) {
  const [display, setDisplay] = useState(0);
  const prevRef = useRef(0);
  const rafRef = useRef(null);

  useEffect(() => {
    const from = prevRef.current;
    const to = value;
    if (from === to) { setDisplay(to); return; }
    const start = performance.now();
    const duration = 650;
    function tick(now) {
      const progress = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - progress, 3);
      setDisplay(from + (to - from) * eased);
      if (progress < 1) {
        rafRef.current = requestAnimationFrame(tick);
      } else {
        prevRef.current = to;
        setDisplay(to);
      }
    }
    if (rafRef.current) cancelAnimationFrame(rafRef.current);
    rafRef.current = requestAnimationFrame(tick);
    return () => { if (rafRef.current) cancelAnimationFrame(rafRef.current); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  return <span className={className}>{formatter(Math.round(display))}</span>;
}

function Avatar({ src, name, size = 40 }) {
  const initials = (name || "?").trim().split(/\s+/).map((w) => w[0]).slice(0, 2).join("").toUpperCase();
  const style = { width: size, height: size, fontSize: Math.max(11, size * 0.35) };
  if (src) {
    return <img src={src} alt={name} style={style} className="rounded-full object-cover shrink-0" />;
  }
  return (
    <div style={style} className="rounded-full bg-[var(--border-input)] text-[var(--text-secondary)] font-semibold flex items-center justify-center shrink-0">
      {initials || <UserIcon size={size * 0.5} />}
    </div>
  );
}

function Shell({ title, userName, avatar, onTitleClick, bottomNav, headerRight, children }) {
  const { accent } = useApp();
  return (
    <div className="min-h-screen bg-[var(--bg-app)] flex flex-col">
      <header className="border-b border-[var(--border-soft)] px-5 py-4 flex items-center justify-between sticky top-0 bg-[var(--bg-app)]/95 backdrop-blur z-10">
        <button
          type="button"
          onClick={onTitleClick}
          disabled={!onTitleClick}
          className={`flex items-center gap-2.5 text-left ${onTitleClick ? "cursor-pointer active:opacity-70" : ""}`}
        >
          {avatar !== undefined ? (
            <Avatar src={avatar} name={userName} size={32} />
          ) : (
            <img src="/logo.svg" alt="" className="w-8 h-8 rounded-lg shrink-0" />
          )}
          <div>
            <div className="text-[var(--text-primary)] font-semibold text-sm leading-tight tracking-tight">{title}</div>
            <div className="text-[var(--text-muted)] text-xs leading-tight">{userName}</div>
          </div>
        </button>
        {headerRight}
      </header>
      <main className={`flex-1 p-5 max-w-4xl w-full mx-auto ${bottomNav ? "pb-32" : ""}`}>{children}</main>
      {bottomNav}
    </div>
  );
}

function TelegramPromptModal({ phase, onLink, onSkip, busy }) {
  return (
    <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-[60] flex items-center justify-center px-6">
      <div className="w-full max-w-sm bg-[var(--bg-panel)] card rounded-2xl p-6 text-center">
        {phase === "success" ? (
          <>
            <div
              className="w-14 h-14 rounded-full flex items-center justify-center mx-auto mb-3"
              style={{ backgroundColor: "var(--good-soft)" }}
            >
              <Check size={28} style={{ color: "var(--good)" }} />
            </div>
            <div className="text-[var(--text-primary)] font-semibold text-base mb-1">Xush kelibsiz!</div>
            <p className="text-[var(--text-secondary)] text-xs">Telegram muvaffaqiyatli ulandi.</p>
          </>
        ) : phase === "waiting" ? (
          <>
            <div
              className="w-14 h-14 rounded-full flex items-center justify-center mx-auto mb-3 animate-pulse"
              style={{ backgroundColor: "rgba(42,169,222,0.15)" }}
            >
              <Send size={24} style={{ color: "#2aa9de" }} />
            </div>
            <div className="text-[var(--text-primary)] font-semibold text-base mb-1.5">Telegramda kuting...</div>
            <p className="text-[var(--text-secondary)] text-xs leading-snug mb-4">
              Ochilgan botga o'ting va "Start" tugmasini bosing — bosishingiz bilan bu oyna avtomatik davom etadi.
            </p>
            <button type="button" onClick={onSkip} className="w-full py-2 text-xs font-medium text-[var(--text-muted)]">
              Keyinroq
            </button>
          </>
        ) : (
          <>
            <div className="flex items-center justify-center gap-2 text-[var(--text-primary)] font-semibold text-base mb-1.5">
              <Send size={18} className="text-[#2aa9de]" /> Xush kelibsiz!
            </div>
            <p className="text-[var(--text-secondary)] text-xs leading-snug mb-4">
              Parolni (yoki loginni) unutib qolsangiz tiklay olishingiz uchun, hisobingizni Telegram botga ulab qo'ying — bir necha soniya vaqt oladi.
            </p>
            <div className="flex flex-col gap-2">
              <button
                type="button"
                disabled={busy}
                onClick={onLink}
                className="w-full py-2.5 rounded-lg text-white text-xs font-semibold disabled:opacity-60"
                style={{ backgroundColor: "#2aa9de" }}
              >
                {busy ? "..." : "Telegramga ulash"}
              </button>
              <button type="button" onClick={onSkip} className="w-full py-2 text-xs font-medium text-[var(--text-muted)]">
                Keyinroq
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

// YANGI: Telegram bot orqali parolni tiklash — ADMIN ham, ISHCHI ham shu bitta
// formadan foydalanadi. Avval login kiritiladi (agar Telegram ulangan bo'lsa,
// botga 6 xonali kod yuboriladi), keyin kod + yangi parol kiritiladi.
function TelegramResetForm({ onBack }) {
  const [step, setStep] = useState("username"); // username -> code -> done
  const [username, setUsername] = useState("");
  const [code, setCode] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);
  const btnShadowRest = "-7px -7px 12px #f8f8f8, 7px 7px 12px #c8c8c8";

  async function submitUsername() {
    setError("");
    if (!username.trim()) { setError("Login kiritilmagan"); return; }
    setBusy(true);
    try {
      const { error: fnErr } = await invokeFn("request-password-reset", { username: username.trim() });
      setBusy(false);
      if (fnErr) { setError(fnErr); return; }
      setInfo("Agar Telegram ulangan bo'lsa, kod shu botga yuborildi. Telegramni tekshiring va kodni kiriting.");
      setStep("code");
    } catch (e) {
      setBusy(false);
      setError(String(e?.message || e));
    }
  }

  async function submitCode() {
    setError("");
    if (!code.trim()) { setError("Kodni kiriting"); return; }
    if (!newPassword || newPassword.length < 6) { setError("Yangi parol kamida 6 belgidan iborat bo'lsin"); return; }
    setBusy(true);
    try {
      const { error: fnErr } = await invokeFn("confirm-password-reset", { username: username.trim(), code: code.trim(), newPassword });
      setBusy(false);
      if (fnErr) { setError(fnErr); return; }
      setStep("done");
    } catch (e) {
      setBusy(false);
      setError(String(e?.message || e));
    }
  }

  if (step === "done") {
    return (
      <div className="space-y-3 text-center">
        <p className="text-sm" style={{ color: "#2f9463" }}>Parol muvaffaqiyatli yangilandi! Endi yangi parol bilan kiring.</p>
        <button type="button" onClick={onBack} className="w-full py-2.5 rounded-lg text-xs font-medium" style={{ background: "#e8e8e8", color: "#4a4a4a", boxShadow: btnShadowRest }}>
          Kirish sahifasiga qaytish
        </button>
      </div>
    );
  }

  if (step === "code") {
    return (
      <div className="space-y-3">
        {info && <p className="text-xs text-center leading-snug" style={{ color: "#6a6a6a" }}>{info}</p>}
        <IconInput icon={<KeyRound size={17} />} value={code} onChange={setCode} placeholder="Telegramdagi 6 xonali kod" />
        <IconInput icon={<Lock size={17} />} type="password" value={newPassword} onChange={setNewPassword} placeholder="Yangi parol" />
        {error && <p className="text-xs text-center" style={{ color: "#a10f0f" }}>{error}</p>}
        <button
          type="button"
          disabled={busy}
          onClick={submitCode}
          className="w-full py-3 rounded-2xl text-sm font-semibold uppercase tracking-widest text-[#f5e9c8] disabled:opacity-60"
          style={{ background: "linear-gradient(155deg, #1a56b0 0%, #123b7a 100%)" }}
        >
          {busy ? "Yuborilmoqda..." : "Parolni tiklash"}
        </button>
        <button type="button" onClick={onBack} className="w-full py-2 text-xs font-medium" style={{ color: "#9a9a9a" }}>Ortga</button>
      </div>
    );
  }

  return (
    <div className="space-y-3">
      <p className="text-center text-sm mb-1 leading-snug" style={{ color: "#6a6a6a" }}>
        Login kiriting — agar Telegram ulangan bo'lsa, tiklash kodi shu yerga yuboriladi.
      </p>
      <IconInput icon={<UserIcon size={17} />} value={username} onChange={setUsername} placeholder="Login" autoFocus />
      <p className="text-center text-[11px] leading-snug" style={{ color: "#9a9a9a" }}>
        Loginingizni ham unutgan bo'lsangiz — Telegramda botga <b>/login</b> deb yozing, u eslatib beradi.
      </p>
      {error && <p className="text-xs text-center" style={{ color: "#a10f0f" }}>{error}</p>}
      <button
        type="button"
        disabled={busy}
        onClick={submitUsername}
        className="w-full py-3 rounded-2xl text-sm font-semibold uppercase tracking-widest text-[#f5e9c8] disabled:opacity-60"
        style={{ background: "linear-gradient(155deg, #1a56b0 0%, #123b7a 100%)" }}
      >
        {busy ? "Yuborilmoqda..." : "Kodni yuborish"}
      </button>
      <button type="button" onClick={onBack} className="w-full py-2 text-xs font-medium" style={{ color: "#9a9a9a" }}>Ortga</button>
    </div>
  );
}

// YANGI: ro'yxatdan o'tishda login band/bo'shligini jonli (debounce bilan) tekshiradi.
// Band bo'lsa — input tagida qizil ogohlantirish; bo'sh bo'lsa — inputning o'ng
// tomonida yashil tick chiqadi.
function UsernameCheckField({ value, onChange, onStatusChange, active }) {
  const [status, setStatus] = useState("idle"); // idle | checking | available | taken | error

  useEffect(() => {
    if (!active) return;
    const v = value.trim();
    if (v.length < 3) {
      setStatus("idle");
      onStatusChange(false);
      return;
    }
    setStatus("checking");
    const handle = setTimeout(async () => {
      try {
        const { data, error } = await supabase.rpc("email_for_username", { p_username: v });
        if (error) { setStatus("idle"); onStatusChange(false); return; }
        if (data) { setStatus("taken"); onStatusChange(false); }
        else { setStatus("available"); onStatusChange(true); }
      } catch (e) {
        setStatus("idle");
        onStatusChange(false);
      }
    }, 450);
    return () => clearTimeout(handle);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, active]);

  return (
    <div>
      <IconInput
        icon={<UserIcon size={17} />}
        value={value}
        onChange={onChange}
        placeholder="Login o'ylab toping"
        autoFocus={active}
        showToggle={status === "available"}
        toggleIcon={<Check size={16} style={{ color: "#2f9463" }} />}
        onToggle={() => {}}
      />
      {status === "taken" && (
        <p className="text-[11px] mt-1.5 pl-1" style={{ color: "#a10f0f" }}>Bu login allaqachon band, boshqasini tanlang</p>
      )}
      {status === "checking" && (
        <p className="text-[11px] mt-1.5 pl-1" style={{ color: "#9a9a9a" }}>Tekshirilmoqda...</p>
      )}
    </div>
  );
}

function LoginScreen({ loginForm, setLoginForm, loginError, onSubmit, onRegister, loginBusy, rememberMe, setRememberMe, lockSeconds, lockTick, loginErrorTick }) {
  const [showPassword, setShowPassword] = useState(false);
  const [registering, setRegistering] = useState(false);
  const [regStep, setRegStep] = useState("username"); // "username" | "password"
  const [usernameAvailable, setUsernameAvailable] = useState(false);
  const [regForm, setRegForm] = useState({ username: "", password: "", confirm: "" });
  const [regError, setRegError] = useState("");
  const [regBusy, setRegBusy] = useState(false);
  const [showRegPassword, setShowRegPassword] = useState(false);
  const [showForgotHint, setShowForgotHint] = useState(false);
  const [forgotMode, setForgotMode] = useState(null); // null | 'choose' | 'employee' | 'admin'
  const [btnHover, setBtnHover] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const [fieldError, setFieldError] = useState(false);
  const [shakeNonce, setShakeNonce] = useState(0);
  const countdownRef = useRef(null);
  const { t } = useApp();

  // YANGI: admin tomondan "bloklandi" signali kelganda (lockTick o'zgarganda) —
  // soniyama-soniya pastga tushadigan sanoqni boshlaymiz va maydonlarni qizil
  // qilib, titratib qo'yamiz.
  useEffect(() => {
    if (!lockTick) return;
    setCountdown(lockSeconds);
    setFieldError(true);
    setShakeNonce((n) => n + 1);
    if (countdownRef.current) clearInterval(countdownRef.current);
    countdownRef.current = setInterval(() => {
      setCountdown((c) => {
        if (c <= 1) {
          clearInterval(countdownRef.current);
          countdownRef.current = null;
          setFieldError(false);
          return 0;
        }
        return c - 1;
      });
    }, 1000);
    return () => { if (countdownRef.current) clearInterval(countdownRef.current); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lockTick]);

  // Oddiy xato (bloklanish emas) — maydonlarni qisqa muddat qizil/titroq qilamiz.
  useEffect(() => {
    if (!loginErrorTick) return;
    setFieldError(true);
    setShakeNonce((n) => n + 1);
    const h = setTimeout(() => setFieldError(false), 2200);
    return () => clearTimeout(h);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loginErrorTick]);

  const locked = countdown > 0;

  async function submitRegister() {
    setRegError("");
    if (regForm.password !== regForm.confirm) {
      setRegError(t("errPasswordMismatch"));
      return;
    }
    setRegBusy(true);
    const result = await onRegister(regForm.username.trim(), regForm.password);
    setRegBusy(false);
    if (result && result.error) {
      setRegError(result.error);
      return;
    }
    // Muvaffaqiyat: parent (WorkforceAppInner) currentUser'ni allaqachon o'rnatgan
    // bo'ladi va ekranni AdminApp'ga almashtiradi; tiklash kodi/Telegram taklifi
    // endi shu yerda emas, dastur darajasida (global overlay sifatida) ko'rsatiladi.
  }

  const cardShadow = "-22px -22px 44px #ffffff, 22px 22px 50px #c3c3c3";
  const titleShadow = "1px 1px 1px rgba(255,255,255,0.9), -2px -2px 1px rgba(163,163,163,0.25)";
  const btnShadowRest = "-7px -7px 12px #f8f8f8, 7px 7px 12px #c8c8c8";
  const btnShadowHover = "0 10px 22px rgba(20,60,140,0.35), -5px -5px 15px rgba(255,255,255,0.6)";

  if (forgotMode) {
    return (
      <div className="min-h-screen flex items-center justify-center px-4" style={{ background: "#e8e8e8" }}>
        <div className="w-full max-w-sm">
          <div className="rounded-[32px] p-8" style={{ background: "#e8e8e8", boxShadow: cardShadow }}>
            <h2 className="text-center text-lg font-bold mb-4" style={{ color: "#4a4a4a" }}>Parolni tiklash</h2>
            <TelegramResetForm onBack={() => setForgotMode(null)} />
          </div>
        </div>
      </div>
    );
  }

  const faceBase = { background: "#e8e8e8", boxShadow: cardShadow, gridArea: "1 / 1", backfaceVisibility: "hidden", WebkitBackfaceVisibility: "hidden" };

  function startRegister() {
    setRegistering(true);
    setRegStep("username");
    setUsernameAvailable(false);
    setRegForm({ username: "", password: "", confirm: "" });
    setRegError("");
  }
  function exitRegister() {
    setRegistering(false);
    setRegStep("username");
    setRegForm({ username: "", password: "", confirm: "" });
    setRegError("");
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-4" style={{ background: "#e8e8e8" }}>
      <div className="w-full max-w-sm" style={{ perspective: "1400px" }}>
        <div
          style={{
            display: "grid",
            transformStyle: "preserve-3d",
            transition: "transform 0.7s cubic-bezier(0.4, 0.1, 0.2, 1)",
            transform: registering ? "rotateY(180deg)" : "rotateY(0deg)",
          }}
        >
          {/* OLD TOMON — Kirish */}
          <div className="rounded-[32px] p-8" style={faceBase}>
            <h1 className="text-center text-3xl font-bold mb-1 tracking-tight" style={{ color: "#4a4a4a", textShadow: titleShadow }}>
              {t("loginHeading")}
            </h1>
            <p className="text-center text-sm mb-7" style={{ color: "#9a9a9a" }}>
              {t("loginSubtitle")}
            </p>

            <div className="space-y-3.5">
              <IconInput
                icon={<UserIcon size={17} />}
                value={loginForm.username}
                onChange={(v) => { setLoginForm({ ...loginForm, username: v }); setFieldError(false); }}
                onKeyDown={(e) => { if (e.key === "Enter") onSubmit(); }}
                placeholder={t("login")}
                autoFocus={!registering}
                disabled={locked}
                error={fieldError}
                shakeKey={shakeNonce}
              />
              <IconInput
                icon={<Lock size={17} />}
                type={showPassword ? "text" : "password"}
                value={loginForm.password}
                onChange={(v) => { setLoginForm({ ...loginForm, password: v }); setFieldError(false); }}
                onKeyDown={(e) => { if (e.key === "Enter") onSubmit(); }}
                placeholder={t("password")}
                showToggle
                toggleIcon={showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                onToggle={() => setShowPassword((v) => !v)}
                disabled={locked}
                error={fieldError}
                shakeKey={shakeNonce}
              />
            </div>

            <div className="flex items-center justify-between mt-4 mb-1">
              <div className="flex items-center gap-2.5">
                <ToggleSwitch checked={rememberMe} onChange={setRememberMe} />
                <span className="text-xs" style={{ color: "#929191" }}>{t("rememberMe")}</span>
              </div>
              <button
                type="button"
                onClick={() => setForgotMode(true)}
                className="text-xs font-medium transition-colors hover:opacity-80"
                style={{ color: "#929191" }}
              >
                {t("forgotPassword")}
              </button>
            </div>

            {(loginError || locked) && (
              <p className="text-xs mt-3 text-center font-medium" style={{ color: "#a10f0f" }}>
                {locked ? `Juda ko'p noto'g'ri urinish. ${countdown} soniyadan keyin qayta urining.` : loginError}
              </p>
            )}

            <button
              type="button"
              disabled={loginBusy || locked}
              onClick={() => onSubmit()}
              onMouseEnter={() => setBtnHover(true)}
              onMouseLeave={() => setBtnHover(false)}
              className="w-full mt-5 py-3.5 rounded-2xl text-sm font-semibold uppercase tracking-widest transition-all duration-300 active:scale-[0.97] disabled:opacity-60"
              style={{
                background: btnHover ? "linear-gradient(155deg, #1a56b0 0%, #123b7a 100%)" : "#e8e8e8",
                color: btnHover ? "#f5e9c8" : "#838383",
                boxShadow: btnHover ? btnShadowHover : btnShadowRest,
                transform: btnHover ? "translateY(-2px)" : "translateY(0)",
              }}
            >
              {locked ? `${countdown}s` : loginBusy ? t("loading") : t("loginBtn")}
            </button>

            <p className="text-center text-xs mt-5" style={{ color: "#9a9a9a" }}>
              {t("noAccountYet")}{" "}
              <button
                type="button"
                onClick={startRegister}
                className="font-semibold transition-opacity hover:opacity-80"
                style={{ color: "#a10f0f" }}
              >
                {t("signUpLink")}
              </button>
            </p>
          </div>

          {/* ORQA TOMON — Ro'yxatdan o'tish (bosqichma-bosqich) */}
          <div className="rounded-[32px] p-8" style={{ ...faceBase, transform: "rotateY(180deg)" }}>
            {regStep === "username" ? (
              <>
                <div className="flex justify-center mb-5">
                  <div
                    className="w-16 h-16 rounded-[22px] flex items-center justify-center"
                    style={{ background: "#e8e8e8", boxShadow: "-6px -6px 12px #f8f8f8, 6px 6px 14px #c3c3c3" }}
                  >
                    <ShieldCheck size={28} style={{ color: "#1a56b0" }} strokeWidth={2} />
                  </div>
                </div>
                <h1 className="text-center text-2xl font-bold mb-1.5 tracking-tight leading-snug" style={{ color: "#4a4a4a", textShadow: titleShadow }}>
                  O'z jamoangizni<br />boshqarishni boshlang
                </h1>
                <p className="text-center text-sm mb-6" style={{ color: "#9a9a9a" }}>{t("registerSubtitle")}</p>

                <UsernameCheckField
                  value={regForm.username}
                  onChange={(v) => setRegForm({ ...regForm, username: v })}
                  onStatusChange={setUsernameAvailable}
                  active={registering && regStep === "username"}
                />

                <button
                  type="button"
                  disabled={!usernameAvailable}
                  onClick={() => setRegStep("password")}
                  className="w-full mt-5 py-3.5 rounded-2xl text-sm font-semibold uppercase tracking-widest transition-all duration-300 active:scale-[0.97] disabled:opacity-40 text-[#f5e9c8]"
                  style={{ background: "linear-gradient(155deg, #1a56b0 0%, #123b7a 100%)" }}
                >
                  Davom etish
                </button>
                <button
                  type="button"
                  onClick={exitRegister}
                  className="w-full mt-3 py-2 rounded-lg text-xs font-medium transition-colors flex items-center justify-center gap-1.5"
                  style={{ color: "#9a9a9a" }}
                >
                  <ArrowLeft size={13} /> {t("alreadyHaveAccount")}
                </button>
              </>
            ) : (
              <>
                <div className="flex justify-center mb-5">
                  <div
                    className="w-16 h-16 rounded-[22px] flex items-center justify-center"
                    style={{ background: "#e8e8e8", boxShadow: "-6px -6px 12px #f8f8f8, 6px 6px 14px #c3c3c3" }}
                  >
                    <Lock size={26} style={{ color: "#1a56b0" }} strokeWidth={2} />
                  </div>
                </div>
                <h1 className="text-center text-2xl font-bold mb-1.5 tracking-tight" style={{ color: "#4a4a4a", textShadow: titleShadow }}>
                  Parol o'ylab toping
                </h1>
                <p className="text-center text-sm mb-6" style={{ color: "#9a9a9a" }}>
                  <b>{regForm.username}</b> uchun parol o'rnating
                </p>

                <div className="space-y-3.5">
                  <IconInput
                    icon={<Lock size={17} />}
                    type={showRegPassword ? "text" : "password"}
                    value={regForm.password}
                    onChange={(v) => setRegForm({ ...regForm, password: v })}
                    placeholder={t("choosePassword")}
                    showToggle
                    toggleIcon={showRegPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                    onToggle={() => setShowRegPassword((v) => !v)}
                    autoFocus
                  />
                  <IconInput
                    icon={<Lock size={17} />}
                    type={showRegPassword ? "text" : "password"}
                    value={regForm.confirm}
                    onChange={(v) => setRegForm({ ...regForm, confirm: v })}
                    placeholder={t("repeatNewPassword")}
                  />
                </div>
                {regError && <p className="text-xs mt-3 text-center" style={{ color: "#a10f0f" }}>{regError}</p>}
                <button
                  type="button"
                  disabled={regBusy}
                  onClick={submitRegister}
                  onMouseEnter={() => setBtnHover(true)}
                  onMouseLeave={() => setBtnHover(false)}
                  className="w-full mt-5 py-3.5 rounded-2xl text-sm font-semibold uppercase tracking-widest transition-all duration-300 active:scale-[0.97] disabled:opacity-60"
                  style={{
                    background: btnHover ? "linear-gradient(155deg, #1a56b0 0%, #123b7a 100%)" : "#e8e8e8",
                    color: btnHover ? "#f5e9c8" : "#838383",
                    boxShadow: btnHover ? btnShadowHover : btnShadowRest,
                    transform: btnHover ? "translateY(-2px)" : "translateY(0)",
                  }}
                >
                  {regBusy ? t("loading") : t("createAccountBtn")}
                </button>
                <button
                  type="button"
                  onClick={() => setRegStep("username")}
                  className="w-full mt-3 py-2 rounded-lg text-xs font-medium transition-colors flex items-center justify-center gap-1.5"
                  style={{ color: "#9a9a9a" }}
                >
                  <ArrowLeft size={13} /> Ortga
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function CopyButton({ text }) {
  const [copied, setCopied] = useState(false);
  const { t } = useApp();
  async function doCopy() {
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch (e) {
    }
  }
  return (
    <button
      type="button"
      onClick={doCopy}
      className="flex items-center gap-1 px-2 py-1 rounded-md field text-[var(--text-secondary)] hover:text-[var(--text-primary)] text-[11px] transition-colors shrink-0"
    >
      {copied ? <Check size={12} className="text-[var(--good)]" /> : <Copy size={12} />}
      {copied ? t("copied") : t("copy")}
    </button>
  );
}

function Lightbox({ src, name, onClose }) {
  return (
    <div
      className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-center justify-center px-6"
      onClick={onClose}
    >
      <button
        type="button"
        onClick={onClose}
        className="absolute top-5 right-5 w-9 h-9 rounded-full bg-black/40 text-white flex items-center justify-center"
        aria-label="Close"
      >
        <X size={18} />
      </button>
      <div onClick={(e) => e.stopPropagation()} className="flex flex-col items-center gap-3">
        {src ? (
          <img src={src} alt={name} className="max-w-[85vw] max-h-[70vh] rounded-2xl object-contain shadow-2xl" />
        ) : (
          <Avatar src={null} name={name} size={180} />
        )}
        <div className="text-white text-sm font-medium">{name}</div>
      </div>
    </div>
  );
}

function isoOf(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// YANGI: "Oylik hisoblash" — Hisobot bo'limidagi ishchi qatoridan ochiladigan, pastdan
// chiquvchi ekran. Tanlangan sanagacha (shu sana ham kiradi) hisob ko'rsatiladi va
// ikkita aniq tugma bor: "To'ladim va yopish" (qoldiq miqdorida ish haqi to'lovi yozadi
// va hisobni yopadi) yoki "Faqat yopish" (qoldiq/qarz keyingi hisobga o'tadi).
function SettleSheet({ emp, previewFor, list, onSettle, onUndo, onClose }) {
  const { accent } = useApp();
  const cur = previewFor(null);
  const cutoff = cur.cutoff;
  const today = todayISO();

  const now = new Date();
  const yest = new Date(now);
  yest.setDate(now.getDate() - 1);
  const eomThis = isoOf(new Date(now.getFullYear(), now.getMonth() + 1, 0));
  const eom = eomThis <= today ? eomThis : isoOf(new Date(now.getFullYear(), now.getMonth(), 0));
  const chips = [
    { label: "Bugun", value: today },
    { label: "Kecha", value: isoOf(yest) },
    { label: "Oy oxiri", value: eom },
  ]
    .filter((c, i, arr) => arr.findIndex((x) => x.value === c.value) === i)
    .filter((c) => !cutoff || c.value > cutoff);

  const [date, setDate] = useState(chips[0]?.value || today);
  const [custom, setCustom] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [showHistory, setShowHistory] = useState(false);
  const [confirmUndo, setConfirmUndo] = useState(false);

  const invalid = date > today || !!(cutoff && date <= cutoff);
  const preview = previewFor(date);
  const rem = preview.remaining;
  const shortDate = (d) => d.slice(5).split("-").reverse().join(".");

  async function run(mode) {
    setBusy(true);
    setError("");
    const r = await onSettle(date, mode);
    setBusy(false);
    if (r && r.error) setError(r.error);
    else onClose();
  }
  async function undo() {
    setBusy(true);
    const r = await onUndo();
    setBusy(false);
    setConfirmUndo(false);
    if (r && r.error) setError(r.error);
  }

  const line = (label, value, tone) => (
    <div className="flex items-center justify-between text-sm py-1">
      <span className="text-[var(--text-muted)]">{label}</span>
      <span className={`font-mono tabular-nums font-semibold ${tone === "bad" ? "text-[var(--bad)]" : tone === "good" ? "text-[var(--good)]" : "text-[var(--text-primary)]"}`}>{value}</span>
    </div>
  );

  return (
    <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-end" onClick={onClose}>
      <div
        className="w-full bg-[var(--bg-panel)] rounded-t-3xl p-5 pb-8 max-h-[88vh] overflow-y-auto"
        style={{ animation: "sheetSlideUp 0.25s ease-out" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="w-10 h-1 rounded-full bg-[var(--border-input)] mx-auto mb-4" />
        <div className="flex items-start justify-between gap-3 mb-1">
          <div className="min-w-0">
            <div className="text-[var(--text-primary)] text-base font-semibold truncate">{emp.name}</div>
            <div className="text-[var(--text-muted)] text-xs mt-0.5">
              {cutoff ? `Oxirgi hisob: ${cutoff}` : "Hali hisoblanmagan"}
            </div>
          </div>
          <button type="button" onClick={onClose} className="text-[var(--text-muted)] hover:text-[var(--text-primary)] shrink-0">
            <X size={18} />
          </button>
        </div>

        <div className="text-[var(--text-secondary)] text-xs font-medium mt-4 mb-2">Qaysi sanagacha hisoblaymiz?</div>
        {chips.length === 0 && !custom ? (
          <p className="text-[var(--warn)] text-xs mb-2">Bugun hisoblangan. Yangi kunlar yozilgandan keyin qayta hisoblang.</p>
        ) : (
          <div className="flex flex-wrap gap-2">
            {chips.map((c) => {
              const active = !custom && date === c.value;
              return (
                <button
                  key={c.value}
                  type="button"
                  onClick={() => { setCustom(false); setDate(c.value); setError(""); }}
                  className="px-3.5 py-2 rounded-lg text-xs font-medium leading-tight text-center transition-colors"
                  style={active ? { backgroundColor: accent, color: "#fff" } : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }}
                >
                  {c.label}
                  <div className="text-[10px] opacity-80 font-mono">{shortDate(c.value)}</div>
                </button>
              );
            })}
            <button
              type="button"
              onClick={() => setCustom(true)}
              className="px-3.5 py-2 rounded-lg text-xs font-medium transition-colors"
              style={custom ? { backgroundColor: accent, color: "#fff" } : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }}
            >
              Boshqa sana
            </button>
          </div>
        )}
        {custom && (
          <div className="mt-3">
            <Field label="Sana (shu sana ham kiradi)" type="date" value={date} onChange={(v) => { setDate(v); setError(""); }} />
          </div>
        )}

        {invalid ? (
          <p className="text-[var(--warn)] text-xs mt-3">
            {date > today ? "Kelajak sanasini tanlab bo'lmaydi." : `Oxirgi hisob ${cutoff} da qilingan, undan keyingi sanani tanlang.`}
          </p>
        ) : (
          <div className="card rounded-xl p-4 mt-4">
            {preview.opening !== 0 && line("Oldingi qoldiq", fmt(preview.opening), preview.opening < 0 ? "bad" : undefined)}
            {line(`Ishlagan (${fmtDays(preview.workedDays)} kun)`, fmt(preview.totalWage))}
            {line("Avans / to'langan", `−${fmt(preview.totalAdvance)}`, "bad")}
            <div className="border-t border-[var(--border-soft)] mt-1.5 pt-2">
              {line(rem < 0 ? "Qarz" : "To'lash kerak", fmt(Math.abs(rem)), rem < 0 ? "bad" : rem > 0 ? "good" : undefined)}
            </div>
          </div>
        )}

        {!invalid && (
          <p className="text-[var(--text-muted)] text-xs mt-3 leading-snug">
            {rem > 0
              ? `"To'ladim va yopish" bosilsa, ${fmt(rem)} ish haqi to'lovi sifatida yoziladi va hisob yopiladi.`
              : rem < 0
                ? `Qarz ${fmt(Math.abs(rem))}. U keyingi hisobga o'tadi.`
                : "Qoldiq yo'q, hisob yopiladi."}
          </p>
        )}

        {error && <p className="text-[var(--bad)] text-xs mt-2">{error}</p>}

        <button
          type="button"
          disabled={busy || invalid}
          onClick={() => run(rem > 0 ? "pay" : "carry")}
          className="w-full mt-4 py-3 rounded-xl text-white text-sm font-semibold hover:opacity-90 transition-opacity disabled:opacity-50"
          style={{ background: accentGradient(accent) }}
        >
          {busy ? "..." : rem > 0 ? "To'ladim va yopish" : "Yopish"}
        </button>
        {rem > 0 && !invalid && (
          <button
            type="button"
            disabled={busy}
            onClick={() => run("carry")}
            className="w-full mt-2 py-2 text-xs font-medium text-[var(--text-secondary)] disabled:opacity-50"
          >
            Faqat yopish (qoldiq keyingi hisobga o'tsin)
          </button>
        )}

        {list.length > 0 && (
          <div className="mt-5 pt-3 border-t border-[var(--border-soft)]">
            <button
              type="button"
              onClick={() => setShowHistory((v) => !v)}
              className="flex items-center gap-1 text-[var(--text-secondary)] text-xs font-medium"
            >
              Oldingi hisoblar ({list.length})
              <ChevronDown size={13} className={`transition-transform ${showHistory ? "rotate-180" : ""}`} />
            </button>
            {showHistory && (
              <div className="mt-2">
                {list.slice().reverse().map((x, i) => (
                  <div key={x.id} className="py-2.5 border-t border-[var(--border-soft)] first:border-t-0">
                    <div className="flex items-center justify-between text-xs">
                      <span className="text-[var(--text-primary)] font-semibold">{x.closedThrough} gacha</span>
                      <span className="text-[var(--text-muted)] font-mono tabular-nums">{fmtDays(x.workedDays)} kun</span>
                    </div>
                    <div className="text-[var(--text-muted)] text-[11px] mt-0.5 font-mono tabular-nums">
                      {fmt(x.totalWage)} − {fmt(x.totalPaid)} = {fmt(x.remaining)}
                    </div>
                    {i === 0 && (
                      !confirmUndo ? (
                        <button type="button" onClick={() => setConfirmUndo(true)} className="text-[var(--bad)] text-[11px] font-medium mt-1">
                          Bekor qilish
                        </button>
                      ) : (
                        <div className="mt-1.5">
                          <p className="text-[var(--bad)] text-[11px] mb-1.5">Shu hisob qayta ochilsinmi? (Yozilgan to'lovlar o'chmaydi)</p>
                          <div className="flex gap-2">
                            <button type="button" disabled={busy} onClick={undo} className="px-3 py-1 rounded-md bg-[var(--bad)] text-white text-[11px] font-semibold">Ha</button>
                            <button type="button" onClick={() => setConfirmUndo(false)} className="px-3 py-1 rounded-md field text-[var(--text-secondary)] text-[11px] font-medium">Yo'q</button>
                          </div>
                        </div>
                      )
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function EmployeeRow({ emp, summary: s, onDelete, onUpdateWage, onResetPassword }) {
  const [open, setOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showPhoto, setShowPhoto] = useState(false);
  const [editingWage, setEditingWage] = useState(false);
  const [wageDraft, setWageDraft] = useState(String(emp.dailyWage || ""));
  const [resetOpen, setResetOpen] = useState(false);
  const [resetPw, setResetPw] = useState("");
  const [resetMsg, setResetMsg] = useState("");
  const { accent, t } = useApp();

  async function saveWage() {
    const n = Number(wageDraft);
    if (wageDraft && n >= 0) await onUpdateWage(n);
    setEditingWage(false);
  }

  async function submitReset() {
    if (!resetPw || resetPw.length < 6) { setResetMsg("Parol kamida 6 belgidan iborat bo'lsin"); return; }
    const result = await onResetPassword(resetPw);
    if (result && result.error) setResetMsg(result.error);
    else { setResetMsg("Yangi parol o'rnatildi!"); setResetPw(""); }
  }

  return (
    <div className="card rounded-xl overflow-hidden">
      {showPhoto && <Lightbox src={emp.avatar} name={emp.name} onClose={() => setShowPhoto(false)} />}
      <div className="p-4 flex items-center justify-between gap-3">
        <div className="flex items-center gap-3 min-w-0">
          <button type="button" onClick={() => setShowPhoto(true)} className="shrink-0" aria-label={t("details")}>
            <Avatar src={emp.avatar} name={emp.name} size={36} />
          </button>
          <div className="min-w-0">
            <div className="text-[var(--text-primary)] text-sm font-medium truncate">{emp.name}</div>
            <div className="text-[var(--text-muted)] text-xs mt-0.5">
              {fmt(emp.dailyWage)}{t("perDay")} &middot; {fmtDays(s.workedDays)} {t("daysWorkedSuffix")}
            </div>
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {s.remaining === 0 ? (
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[10px] font-semibold" style={{ backgroundColor: "var(--warn-soft)", color: "var(--warn)" }}>
              <Check size={11} /> To'liq to'landi
            </span>
          ) : (
            <div className={`text-sm font-semibold font-mono tabular-nums ${s.remaining < 0 ? "text-[var(--bad)]" : "text-[var(--good)]"}`}>
              {fmt(s.remaining)}
            </div>
          )}
          <button
            type="button"
            onClick={() => setOpen((v) => !v)}
            className={`p-1.5 rounded-md transition-colors ${open ? "text-[#12161c]" : "text-[var(--text-muted)] hover:text-[var(--text-primary)]"}`}
            style={open ? { backgroundColor: accent } : undefined}
            aria-label={t("details")}
          >
            <MoreVertical size={16} />
          </button>
        </div>
      </div>


      {open && (
        <div className="border-t border-[var(--border)] p-4 space-y-1">
          <button type="button" onClick={() => setShowPhoto(true)} className="w-full flex flex-col items-center gap-2 pb-3">
            <Avatar src={emp.avatar} name={emp.name} size={64} />
            <div className="text-[var(--text-primary)] text-sm font-semibold">{emp.name}</div>
          </button>

          <div className="flex items-center justify-between gap-2 py-2.5 border-b border-[var(--border-soft)]">
            <div className="min-w-0">
              <div className="flex items-center gap-1.5 text-[var(--text-muted)] text-[11px] mb-0.5"><KeyRound size={11} /> {t("login")}</div>
              <div className="text-[var(--text-primary)] text-sm truncate">{emp.username}</div>
            </div>
            <CopyButton text={emp.username} />
          </div>

          {!resetOpen ? (
            <button
              type="button"
              onClick={() => { setResetOpen(true); setResetMsg(""); }}
              className="w-full flex items-center gap-1.5 py-2.5 border-b border-[var(--border-soft)] text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors"
            >
              <KeyRound size={13} /> Parolni tiklash
            </button>
          ) : (
            <div className="field rounded-lg p-3 space-y-2 my-2">
              <Field label="Yangi parol" type="password" value={resetPw} onChange={setResetPw} />
              {resetMsg && <p className="text-[var(--text-secondary)] text-xs">{resetMsg}</p>}
              <div className="flex gap-2">
                <button type="button" onClick={submitReset} className="flex-1 py-2 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity" style={{ background: accentGradient(accent) }}>
                  {t("save")}
                </button>
                <button type="button" onClick={() => { setResetOpen(false); setResetPw(""); }} className="flex-1 py-2 rounded-lg field text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors">
                  {t("cancel")}
                </button>
              </div>
            </div>
          )}

          {!editingWage ? (
            <div className="flex items-center justify-between gap-2 py-2.5 border-b border-[var(--border-soft)]">
              <div className="min-w-0">
                <div className="flex items-center gap-1.5 text-[var(--text-muted)] text-[11px] mb-0.5"><Wallet size={11} /> {t("dailyWage")}</div>
                <div className="text-[var(--text-primary)] text-sm font-mono tabular-nums">{fmt(emp.dailyWage)}{t("perDay")}</div>
              </div>
              <button
                type="button"
                onClick={() => { setWageDraft(String(emp.dailyWage || "")); setEditingWage(true); }}
                className="p-1.5 rounded-md text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors shrink-0"
                aria-label={t("editWage")}
              >
                <Settings size={14} />
              </button>
            </div>
          ) : (
            <div className="field rounded-lg p-3 space-y-2 my-2">
              <MoneyField label={t("newDailyWage")} value={wageDraft} onChange={setWageDraft} suffix="so'm" />
              <div className="flex gap-2">
                <button type="button" onClick={saveWage} className="flex-1 py-2 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity" style={{ background: accentGradient(accent) }}>
                  {t("save")}
                </button>
                <button type="button" onClick={() => setEditingWage(false)} className="flex-1 py-2 rounded-lg field text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors">
                  {t("cancel")}
                </button>
              </div>
            </div>
          )}

          <div className="grid grid-cols-3 gap-2 pt-3 pb-1">
            <div className="text-center">
              <div className="text-[var(--text-primary)] text-sm font-semibold font-mono tabular-nums">{fmtDays(s.workedDays)}</div>
              <div className="text-[9px] text-[var(--text-muted)] uppercase tracking-wide mt-0.5">{t("day")}</div>
            </div>
            <div className="text-center">
              <div className="text-[var(--bad)] text-sm font-semibold font-mono tabular-nums">{fmt(s.totalAdvance)}</div>
              <div className="text-[9px] text-[var(--text-muted)] uppercase tracking-wide mt-0.5">{t("advance")}</div>
            </div>
            <div className="text-center">
              {s.remaining === 0 ? (
                <div className="flex items-center justify-center gap-1 text-[var(--warn)] text-sm font-semibold">
                  <Check size={13} />
                </div>
              ) : (
                <div className="text-[var(--good)] text-sm font-semibold font-mono tabular-nums">{fmt(s.remaining)}</div>
              )}
              <div className="text-[9px] text-[var(--text-muted)] uppercase tracking-wide mt-0.5">{t("remaining")}</div>
            </div>
          </div>

          {!confirmDelete ? (
            <button
              type="button"
              onClick={() => setConfirmDelete(true)}
              className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg bg-[var(--bad-soft)] text-[var(--bad)] text-xs font-medium hover:opacity-90 transition-opacity mt-1"
            >
              <Trash2 size={13} /> {t("deleteEmployee")}
            </button>
          ) : (
            <div className="mt-1 bg-[var(--bad-soft)] border border-[var(--bad)]/30 rounded-lg p-3">
              <p className="text-[var(--bad)] text-xs mb-2.5">{t("confirmDelete", { name: emp.name })}</p>
              <div className="flex gap-2">
                <button type="button" onClick={onDelete} className="flex-1 py-2 rounded-lg bg-[var(--bad)] text-white text-xs font-semibold hover:opacity-90 transition-opacity">
                  {t("yesDelete")}
                </button>
                <button type="button" onClick={() => setConfirmDelete(false)} className="flex-1 py-2 rounded-lg field text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors">
                  {t("cancel")}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// YANGI: davomat holatini bosganda rang darhol "qattiq" almashmasin deb,
// qisqa "pop" (kattalashib-qaytish) animatsiyasi qo'shildi.
function AttendanceStatusRow({ emp, status, isFuture, locked, onCycle }) {
  const { t } = useApp();
  const [pulse, setPulse] = useState(false);
  const pulseTimer = useRef(null);

  useEffect(() => () => { if (pulseTimer.current) clearTimeout(pulseTimer.current); }, []);

  function handleClick() {
    if (isFuture || locked) return;
    onCycle();
    setPulse(false);
    requestAnimationFrame(() => {
      setPulse(true);
      if (pulseTimer.current) clearTimeout(pulseTimer.current);
      pulseTimer.current = setTimeout(() => setPulse(false), 280);
    });
  }

  const statusConfig = {
    null: { label: t("statusNone"), icon: <Calendar size={15} />, bg: "var(--bg-app)", color: "var(--text-muted)", border: "1px solid var(--border-input)" },
    1: { label: t("fullDay"), icon: <CheckCircle2 size={15} />, bg: "var(--good)", color: "#0e1712", border: "none" },
    0.5: { label: t("halfDay"), icon: <Calendar size={15} />, bg: "var(--warn)", color: "#1a1608", border: "none" },
    0: { label: t("absent"), icon: <XCircle size={15} />, bg: "var(--bad)", color: "#1c0e0c", border: "none" },
  };
  const cfg = statusConfig[status === null ? "null" : status];

  return (
    <button
      type="button"
      disabled={isFuture || locked}
      onClick={handleClick}
      className="w-full card rounded-xl p-3.5 flex items-center justify-between gap-3 disabled:opacity-40 disabled:cursor-not-allowed transition-colors text-left"
    >
      <div className="flex items-center gap-2.5 min-w-0">
        <Avatar src={emp.avatar} name={emp.name} size={32} />
        <div className="min-w-0">
          <div className="text-[var(--text-primary)] text-sm font-medium truncate">{emp.name}</div>
          <div className="text-[var(--text-muted)] text-[11px]">
            {locked ? "Hisoblangan davr" : `${fmt(emp.dailyWage)}${t("perDay")}`}
          </div>
        </div>
      </div>
      <span
        className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold shrink-0 ${pulse ? "status-pop" : ""}`}
        style={{ backgroundColor: cfg.bg, color: cfg.color, border: cfg.border }}
      >
        {cfg.icon} {cfg.label}
      </span>
    </button>
  );
}

function MenuRow({ icon, label, onClick, danger }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="w-full flex items-center justify-between py-3.5 border-b border-[var(--border)] last:border-b-0 text-left"
    >
      <span className={`flex items-center gap-3 text-sm ${danger ? "text-[var(--bad)]" : "text-[var(--text-primary)]"}`}>{icon}{label}</span>
      <ChevronRight size={16} className="text-[var(--text-muted)]" />
    </button>
  );
}

function NotificationPanel({ open, onClose, notifications, onMarkAllRead, onMarkRead }) {
  const { t } = useApp();
  const [selected, setSelected] = useState(null);
  const hasUnread = notifications.some((n) => !n.is_read);

  function openNotification(n) {
    setSelected(n);
    if (!n.is_read) onMarkRead(n.id);
  }

  return (
    <>
      {open && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-30 transition-opacity" onClick={onClose} />
      )}
      <div
        className={`fixed top-0 right-0 h-full w-[86%] max-w-sm bg-[var(--bg-panel)] card z-40 overflow-y-auto transition-transform duration-200 ${
          open ? "translate-x-0" : "translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--border)] sticky top-0 bg-[var(--bg-panel)] z-10">
          <span className="text-[var(--text-primary)] text-sm font-semibold flex items-center gap-1.5">
            <Bell size={16} /> {t("notifications")}
          </span>
          <div className="flex items-center gap-1">
            {hasUnread && (
              <button
                type="button"
                onClick={onMarkAllRead}
                aria-label={t("markAllRead")}
                title={t("markAllRead")}
                className="p-1.5 rounded-md text-[var(--accent)] hover:bg-[var(--bg-app)] transition-colors"
              >
                <CheckCheck size={17} />
              </button>
            )}
            <button type="button" onClick={onClose} className="p-1.5 text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors">
              <X size={18} />
            </button>
          </div>
        </div>

        <div className="px-5 py-4 space-y-2.5">
          {notifications.length === 0 && (
            <p className="text-[var(--text-muted)] text-sm text-center py-10">{t("noNotifications")}</p>
          )}
          {notifications.map((n) => (
            <button
              type="button"
              key={n.id}
              onClick={() => openNotification(n)}
              className="w-full text-left card rounded-xl p-3.5 active:scale-[0.99] transition-transform"
              style={!n.is_read ? { borderColor: "var(--accent)" } : undefined}
            >
              <div className="flex items-center justify-between gap-2 mb-1">
                <span className="text-[var(--text-primary)] text-sm font-semibold">{n.title}</span>
                {!n.is_read && <span className="w-2 h-2 rounded-full shrink-0" style={{ backgroundColor: "var(--accent)" }} />}
              </div>
              <p className="text-[var(--text-secondary)] text-xs leading-snug line-clamp-2">{n.body}</p>
              <p className="text-[var(--text-faint)] text-[10px] mt-1.5">
                {new Date(n.created_at).toLocaleString()}
              </p>
            </button>
          ))}
        </div>
      </div>

      {/* To'liq xabar — pastdan chiqadigan varaq */}
      {selected && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-end" onClick={() => setSelected(null)}>
          <div
            className="w-full bg-[var(--bg-panel)] rounded-t-3xl p-5 pb-8 max-h-[70vh] overflow-y-auto"
            style={{ animation: "sheetSlideUp 0.25s ease-out" }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="w-10 h-1 rounded-full bg-[var(--border-input)] mx-auto mb-4" />
            <div className="flex items-start justify-between gap-3 mb-2">
              <span className="text-[var(--text-primary)] text-base font-semibold">{selected.title}</span>
              <button
                type="button"
                onClick={() => setSelected(null)}
                className="text-[var(--text-muted)] hover:text-[var(--text-primary)] shrink-0"
              >
                <X size={18} />
              </button>
            </div>
            <p className="text-[var(--text-faint)] text-[11px] mb-3">
              {new Date(selected.created_at).toLocaleString()}
            </p>
            <p className="text-[var(--text-secondary)] text-sm leading-relaxed whitespace-pre-wrap">
              {selected.body}
            </p>
          </div>
        </div>
      )}
    </>
  );
}

// YANGI: ma'lumotlarni yuklab olish — Instagram'dagi "Download your information"
// ekraniga o'xshab, nima yuklanishini ro'yxat qilib ko'rsatadi, "Tayyorlanmoqda..."
// bosqichidan o'tadi, so'ng muvaffaqiyatli yakunlanganini bildiradi.
function DataExportSheet({ open, onClose, onExport }) {
  const { accent } = useApp();
  const [phase, setPhase] = useState("idle"); // idle | preparing | done

  useEffect(() => {
    if (!open) setPhase("idle");
  }, [open]);

  async function start() {
    setPhase("preparing");
    await new Promise((r) => setTimeout(r, 900));
    onExport();
    setPhase("done");
    setTimeout(() => onClose(), 1400);
  }

  if (!open) return null;
  const items = [
    { icon: <Users size={16} />, label: "Ishchilar" },
    { icon: <Calendar size={16} />, label: "Davomat" },
    { icon: <Wallet size={16} />, label: "Avanslar" },
    { icon: <ClipboardList size={16} />, label: "Hisob-kitoblar" },
  ];

  return (
    <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-end" onClick={phase === "idle" ? onClose : undefined}>
      <div
        className="w-full bg-[var(--bg-panel)] rounded-t-3xl p-5 pb-8"
        style={{ animation: "sheetSlideUp 0.25s ease-out" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="w-10 h-1 rounded-full bg-[var(--border-input)] mx-auto mb-4" />
        {phase === "done" ? (
          <div className="flex flex-col items-center gap-3 py-6 text-center">
            <div className="w-14 h-14 rounded-full flex items-center justify-center" style={{ backgroundColor: "var(--good-soft)" }}>
              <Check size={28} style={{ color: "var(--good)" }} />
            </div>
            <div className="text-[var(--text-primary)] font-semibold text-base">Fayl tayyor</div>
            <p className="text-[var(--text-secondary)] text-xs">Ma'lumotlar yuklab olindi.</p>
          </div>
        ) : (
          <>
            <div className="text-[var(--text-primary)] text-base font-semibold text-center mb-1">Ma'lumotlarni yuklab olish</div>
            <p className="text-[var(--text-muted)] text-xs text-center mb-5 leading-snug">
              Quyidagilarning barchasi bitta faylga saqlanadi:
            </p>
            <div className="space-y-2 mb-5">
              {items.map((it) => (
                <div key={it.label} className="flex items-center gap-3 field rounded-lg px-3.5 py-2.5">
                  <span style={{ color: accent }}>{it.icon}</span>
                  <span className="text-[var(--text-primary)] text-sm">{it.label}</span>
                </div>
              ))}
            </div>
            <button
              type="button"
              disabled={phase === "preparing"}
              onClick={start}
              className="w-full py-3 rounded-xl text-white text-sm font-semibold hover:opacity-90 transition-opacity disabled:opacity-70"
              style={{ background: accentGradient(accent) }}
            >
              {phase === "preparing" ? "Tayyorlanmoqda..." : "Yuklab olish"}
            </button>
            {phase === "idle" && (
              <button type="button" onClick={onClose} className="w-full mt-2 py-2 text-xs font-medium text-[var(--text-secondary)]">
                Bekor qilish
              </button>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function ProfileDrawer({
  open, onClose, me, roleLabel, isAdmin, onDeleteAccount, onLogout, onExportBackup,
  changeOwnCredentials, updateAvatar, enableNotifications, linkTelegram,
  accent, setAccent, mode, setMode, fontScale, setFontScale, lang, setLang,
}) {
  const [tgBusy, setTgBusy] = useState(false);
  const [tgMsg, setTgMsg] = useState("");
  async function handleLinkTelegram() {
    setTgBusy(true);
    setTgMsg("");
    const result = await linkTelegram();
    setTgBusy(false);
    if (result && result.error) setTgMsg(result.error);
    else setTgMsg("Telegram ochildi — u yerda \"Start\" tugmasini bosing.");
  }
  const { t } = useApp();
  const fileRef = useRef(null);
  const [page, setPage] = useState(null);
  const [currentPw, setCurrentPw] = useState("");
  const [newUsername, setNewUsername] = useState(me.username);
  const [newPw, setNewPw] = useState("");
  const [newPw2, setNewPw2] = useState("");
  const [msg, setMsg] = useState({ type: "", text: "" });
  const [avatarBusy, setAvatarBusy] = useState(false);
  const [saveBusy, setSaveBusy] = useState(false);
  const [confirmDeleteAcc, setConfirmDeleteAcc] = useState(false);
  const [deletePwInput, setDeletePwInput] = useState("");
  const [deletePwError, setDeletePwError] = useState("");
  const [deleteBusy, setDeleteBusy] = useState(false);
  const [exportSheetOpen, setExportSheetOpen] = useState(false);

  useEffect(() => {
    if (!open) setPage(null);
  }, [open]);

  const PAGE_TITLES = { appearance: t("appearance"), privacy: t("privacySecurity"), credentials: t("updateCredentials"), language: t("language"), advanced: t("advanced") };
  const PARENT_PAGE = { credentials: "privacy" };
  const breadcrumbTrail = (() => {
    const trail = [];
    let cur = page;
    while (cur) {
      trail.unshift(cur);
      cur = PARENT_PAGE[cur];
    }
    return trail;
  })();

  async function handleFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setAvatarBusy(true);
    try {
      const dataUrl = await fileToAvatarDataUrl(file);
      await updateAvatar(dataUrl);
    } catch (err) {
    }
    setAvatarBusy(false);
    if (fileRef.current) fileRef.current.value = "";
  }

  async function submitPassword() {
    setMsg({ type: "", text: "" });
    if (!newUsername.trim()) {
      setMsg({ type: "error", text: t("errEmptyLogin") });
      return;
    }
    const wantsPasswordChange = !!newPw || !!newPw2;
    if (wantsPasswordChange && (!newPw || newPw.length < 6)) {
      setMsg({ type: "error", text: "Yangi parol kamida 6 belgidan iborat bo'lsin" });
      return;
    }
    if (wantsPasswordChange && newPw !== newPw2) {
      setMsg({ type: "error", text: t("errPasswordMismatch") });
      return;
    }
    setSaveBusy(true);
    const result = await changeOwnCredentials(newUsername.trim(), wantsPasswordChange ? newPw : null, currentPw);
    setSaveBusy(false);
    if (result && result.error) {
      setMsg({ type: "error", text: result.error });
      return;
    }
    setCurrentPw(""); setNewPw(""); setNewPw2("");
    setMsg({ type: "ok", text: t("savedOk") });
  }

  const currentFontLabel = t(FONT_SCALES.find((f) => f.value === fontScale)?.key || "fontMedium");

  return (
    <>
      {open && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-30 transition-opacity" onClick={onClose} />
      )}
      <div
        className={`fixed top-0 left-0 h-full w-[86%] max-w-sm bg-[var(--bg-panel)] card z-40 overflow-y-auto transition-transform duration-200 ${
          open ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-[var(--border)] sticky top-0 bg-[var(--bg-panel)] z-10">
          {page ? (
            <button
              type="button"
              onClick={() => setPage(PARENT_PAGE[page] || null)}
              className="flex items-center gap-1.5 text-[var(--text-primary)] text-sm font-semibold"
            >
              <ArrowLeft size={18} /> {PAGE_TITLES[page]}
            </button>
          ) : <span />}
          <button
            type="button"
            onClick={() => setMode(mode === "dark" ? "light" : "dark")}
            className="text-[var(--text-secondary)] hover:text-[var(--accent)] transition-colors"
            aria-label={t("themeColor")}
          >
            {mode === "dark" ? <Sun size={20} /> : <Moon size={20} />}
          </button>
        </div>

        {page && (
          <div className="flex items-center gap-1 px-5 pt-2.5 pb-1 text-[11px] flex-wrap">
            <button
              type="button"
              onClick={() => setPage(null)}
              className="text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors"
            >
              {t("profile")}
            </button>
            {breadcrumbTrail.map((p, i) => {
              const isLast = i === breadcrumbTrail.length - 1;
              return (
                <span key={p} className="flex items-center gap-1">
                  <ChevronRight size={11} className="text-[var(--text-faint)]" />
                  <button
                    type="button"
                    onClick={() => !isLast && setPage(p)}
                    disabled={isLast}
                    className={isLast ? "text-[var(--text-primary)] font-medium" : "text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors"}
                  >
                    {PAGE_TITLES[p]}
                  </button>
                </span>
              );
            })}
          </div>
        )}

        {!page && (
          <>
            <div className="px-5 pt-5">
              <div className="flex items-center gap-3.5 pb-5">
                <div className="relative shrink-0">
                  <Avatar src={me.avatar} name={me.name} size={56} />
                  <button
                    type="button"
                    onClick={() => fileRef.current && fileRef.current.click()}
                    disabled={avatarBusy}
                    className="absolute -bottom-1 -right-1 w-6 h-6 rounded-full flex items-center justify-center border-2"
                    style={{ backgroundColor: accent, color: "#12161c", borderColor: "var(--bg-panel)" }}
                    aria-label={t("changePhoto")}
                  >
                    <Camera size={11} />
                  </button>
                  <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={handleFile} />
                </div>
                <div className="min-w-0">
                  <div className="text-[var(--text-primary)] text-base font-semibold truncate">{me.name}</div>
                  <div className="text-[var(--text-muted)] text-sm truncate">{roleLabel}</div>
                </div>
              </div>
            </div>

            <div className="px-5">
              <MenuRow icon={<Paintbrush size={18} className="text-[var(--accent)]" />} label={t("appearance")} onClick={() => setPage("appearance")} />
              <MenuRow icon={<ShieldCheck size={18} className="text-[var(--good)]" />} label={t("privacySecurity")} onClick={() => setPage("privacy")} />
              <MenuRow icon={<Globe size={18} className="text-[var(--accent)]" />} label={t("language")} onClick={() => setPage("language")} />
              <MenuRow icon={<UserX size={18} className="text-[var(--bad)]" />} label={t("advanced")} onClick={() => setPage("advanced")} />
              <MenuRow icon={<LogOut size={18} className="text-[var(--bad)]" />} label={t("logout")} onClick={onLogout} danger />
            </div>
            <div className="h-5" />
          </>
        )}

        {page === "appearance" && (
          <div className="px-5 pt-5 pb-8 space-y-6">
            <div>
              <div className="flex items-center gap-1.5 text-[var(--text-secondary)] text-xs font-medium mb-2.5">
                <Palette size={13} /> {t("themeColor")}
              </div>
              <div className="flex flex-wrap gap-2.5">
                {ACCENT_PRESETS.map((p) => (
                  <button
                    key={p.value}
                    type="button"
                    onClick={() => setAccent(p.value)}
                    className="w-9 h-9 rounded-full flex items-center justify-center transition-transform active:scale-90"
                    style={{ backgroundColor: p.value, boxShadow: accent === p.value ? `0 0 0 2px var(--bg-panel), 0 0 0 4px ${p.value}` : "none" }}
                    aria-label={p.name}
                  >
                    {accent === p.value && <Check size={14} className="text-[#12161c]" />}
                  </button>
                ))}
              </div>
            </div>
            <div>
              <div className="flex items-center gap-1.5 text-[var(--text-secondary)] text-xs font-medium mb-2.5">
                <Type size={13} /> {t("fontSize")}
              </div>
              <div className="flex gap-2">
                {FONT_SCALES.map((f) => (
                  <button
                    key={f.value}
                    type="button"
                    onClick={() => setFontScale(f.value)}
                    className="flex-1 py-2.5 rounded-lg text-xs font-medium transition-colors"
                    style={
                      fontScale === f.value
                        ? { backgroundColor: accent, color: "#12161c" }
                        : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }
                    }
                  >
                    {t(f.key)}
                  </button>
                ))}
              </div>
            </div>
          </div>
        )}

        {page === "privacy" && (
          <div className="px-5">
            <MenuRow icon={<KeyRound size={18} className="text-[var(--accent)]" />} label={t("updateCredentials")} onClick={() => setPage("credentials")} />
            <MenuRow icon={<Send size={18} className="text-[#2aa9de]" />} label={tgBusy ? t("loading") : "Telegramga ulash (parolni tiklash uchun)"} onClick={handleLinkTelegram} />
            {tgMsg && <p className="text-[var(--text-secondary)] text-xs pb-3 -mt-1">{tgMsg}</p>}
            <MenuRow icon={<Send size={18} className="text-[var(--good)]" />} label={t("enableNotifications")} onClick={enableNotifications} />
          </div>
        )}

        {page === "credentials" && (
          <div className="px-5 pt-5 pb-8">
            <div className="space-y-3">
              <Field label={t("currentPassword")} type="password" value={currentPw} onChange={setCurrentPw} />
              <Field label={t("newLogin")} value={newUsername} onChange={setNewUsername} />
              <Field label={t("newPassword") + " (ixtiyoriy)"} type="password" value={newPw} onChange={setNewPw} />
              <Field label={t("repeatNewPassword")} type="password" value={newPw2} onChange={setNewPw2} />
            </div>
            {msg.text && (
              <p className={`text-xs mt-2.5 ${msg.type === "error" ? "text-[var(--bad)]" : "text-[var(--good)]"}`}>{msg.text}</p>
            )}
            <button
              type="button"
              disabled={saveBusy}
              onClick={submitPassword}
              className="mt-3 w-full py-2.5 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity disabled:opacity-60"
              style={{ background: accentGradient(accent) }}
            >
              {saveBusy ? t("loading") : t("save")}
            </button>
          </div>
        )}

        {page === "language" && (
          <div className="px-5 pt-5 pb-8">
            <div className="flex flex-col gap-2">
              {LANGS.map((l) => (
                <button
                  key={l.code}
                  type="button"
                  onClick={() => setLang(l.code)}
                  className="flex items-center justify-between px-3.5 py-2.5 rounded-lg text-sm transition-colors"
                  style={
                    lang === l.code
                      ? { backgroundColor: accent, color: "#12161c", fontWeight: 600 }
                      : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }
                  }
                >
                  {l.label}
                  {lang === l.code && <Check size={14} />}
                </button>
              ))}
            </div>
          </div>
        )}

        {page === "advanced" && (
          <div className="px-5 pt-5 pb-8">
            {isAdmin && onExportBackup && (
              <>
                <button
                  type="button"
                  onClick={() => setExportSheetOpen(true)}
                  className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded-lg field text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors"
                >
                  <Download size={14} /> Ma'lumotlarni yuklab olish
                </button>
                <p className="text-[var(--text-faint)] text-[11px] mt-1.5 mb-5 leading-snug">
                  Barcha ishchilar, davomat, avanslar va hisob-kitoblar bitta faylga saqlanadi.
                </p>
              </>
            )}
            {!confirmDeleteAcc ? (
              <button
                type="button"
                onClick={() => setConfirmDeleteAcc(true)}
                className="w-full flex items-center justify-center gap-1.5 py-2.5 rounded-lg bg-[var(--bad-soft)] text-[var(--bad)] text-xs font-medium hover:opacity-90 transition-opacity"
              >
                <UserX size={14} /> {t("deleteAccount")}
              </button>
            ) : (
              <div className="bg-[var(--bad-soft)] border border-[var(--bad)]/30 rounded-lg p-3.5">
                <p className="text-[var(--bad)] text-xs mb-2.5">
                  {isAdmin ? t("confirmDeleteAccountAdmin") : t("confirmDeleteAccountEmployee")}
                </p>
                <div className="mb-2.5">
                  <Field label={t("currentPassword")} type="password" value={deletePwInput} onChange={setDeletePwInput} />
                  {deletePwError && <p className="text-[var(--bad)] text-xs mt-1.5">{deletePwError}</p>}
                </div>
                <div className="flex gap-2">
                  <button
                    type="button"
                    disabled={deleteBusy}
                    onClick={async () => {
                      setDeleteBusy(true);
                      const result = await onDeleteAccount(deletePwInput);
                      setDeleteBusy(false);
                      if (result && result.error) {
                        setDeletePwError(result.error);
                      }
                    }}
                    className="flex-1 py-2 rounded-lg bg-[var(--bad)] text-white text-xs font-semibold hover:opacity-90 transition-opacity disabled:opacity-60"
                  >
                    {deleteBusy ? t("loading") : t("yesDeleteAccount")}
                  </button>
                  <button
                    type="button"
                    onClick={() => { setConfirmDeleteAcc(false); setDeletePwInput(""); setDeletePwError(""); }}
                    className="flex-1 py-2 rounded-lg field text-[var(--text-secondary)] text-xs font-medium hover:text-[var(--text-primary)] transition-colors"
                  >
                    {t("cancel")}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
      {isAdmin && onExportBackup && (
        <DataExportSheet
          open={exportSheetOpen}
          onClose={() => setExportSheetOpen(false)}
          onExport={onExportBackup}
        />
      )}
    </>
  );
}

// YANGI: Avanslar tarixi endi oylarga guruhlangan holda (akkordeon) ko'rsatiladi —
// ro'yxat cheksiz cho'zilib ketmasligi uchun. Har bir oy yopiq/ochiq bo'lishi mumkin,
// sarlavhada shu oy uchun jami summa ham ko'rsatiladi. Yozuvni o'chirishdan oldin
// endi "Ha / Yo'q" bilan tasdiqlash so'raladi — tasodifiy bosilib ketishning oldini oladi.
function AdvanceHistoryRow({ a, t, onDelete }) {
  const [confirming, setConfirming] = useState(false);
  return (
    <div className="flex items-center gap-3 py-2.5 border-b border-[var(--border-soft)] last:border-b-0">
      <div
        className="w-8 h-8 rounded-full flex items-center justify-center shrink-0"
        style={a.type === "salary" ? { backgroundColor: "var(--good-soft)", color: "var(--good)" } : { backgroundColor: "var(--warn-soft)", color: "var(--warn)" }}
      >
        {a.type === "salary" ? <Wallet size={14} /> : <TrendingDown size={14} />}
      </div>
      <div className="min-w-0 flex-1">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-[var(--text-primary)] text-sm font-semibold font-mono tabular-nums">{fmt(a.amount)}</span>
          <span
            className="text-[9px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full"
            style={a.type === "salary" ? { backgroundColor: "var(--good-soft)", color: "var(--good)" } : { backgroundColor: "var(--warn-soft)", color: "var(--warn)" }}
          >
            {a.type === "salary" ? t("typeSalary") : t("typeAvans")}
          </span>
        </div>
        <div className="text-xs mt-0.5 truncate">
          <span className="text-[var(--text-muted)]">{a.date}</span>
          {a.note && <span className="text-[var(--text-primary)] font-semibold"> · {a.note}</span>}
        </div>
      </div>

      {!confirming ? (
        <button
          type="button"
          onClick={() => setConfirming(true)}
          className="p-1.5 rounded-md text-[var(--text-muted)] hover:text-[var(--bad)] hover:bg-[var(--bad-soft)] transition-colors shrink-0"
          aria-label="O'chirish"
        >
          <Trash2 size={14} />
        </button>
      ) : (
        <div className="flex items-center gap-1 shrink-0">
          <button
            type="button"
            onClick={() => { onDelete(); setConfirming(false); }}
            className="px-2.5 py-1 rounded-md bg-[var(--bad)] text-white text-[10px] font-semibold"
          >
            Ha
          </button>
          <button
            type="button"
            onClick={() => setConfirming(false)}
            className="px-2.5 py-1 rounded-md field text-[var(--text-secondary)] text-[10px] font-medium"
          >
            Yo'q
          </button>
        </div>
      )}
    </div>
  );
}

function AdvanceHistoryCard({ list, onDelete }) {
  const { t, lang } = useApp();
  const localeTag = lang === "ru" ? "ru-RU" : lang === "en" ? "en-US" : "uz-UZ";
  const monthNamesUz = ["Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun", "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"];

  const sorted = list.slice().sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

  const groups = [];
  const groupIndex = {};
  for (const a of sorted) {
    const key = a.date.slice(0, 7);
    if (!(key in groupIndex)) {
      groupIndex[key] = groups.length;
      groups.push({ key, items: [] });
    }
    groups[groupIndex[key]].items.push(a);
  }

  const [openKey, setOpenKey] = useState(groups[0]?.key ?? null);

  function monthLabel(key) {
    const [y, m] = key.split("-").map(Number);
    const name = lang === "uz" ? monthNamesUz[m - 1] : new Date(y, m - 1, 1).toLocaleDateString(localeTag, { month: "long" });
    return `${name} ${y}`;
  }

  return (
    <div className="card rounded-xl p-5">
      <div className="text-[var(--text-primary)] text-sm font-semibold mb-3">{t("advanceHistory")}</div>

      {list.length === 0 && (
        <div className="flex flex-col items-center gap-2 py-6 text-center">
          <Wallet size={18} className="text-[var(--text-faint)]" />
          <p className="text-[var(--text-muted)] text-xs">{t("noAdvances")}</p>
        </div>
      )}

      <div>
        {groups.map((g, idx) => {
          const isOpen = openKey === g.key;
          const groupTotal = g.items.reduce((sum, a) => sum + Number(a.amount), 0);
          return (
            <div key={g.key} className={idx > 0 ? "border-t border-[var(--border-soft)]" : ""}>
              <button
                type="button"
                onClick={() => setOpenKey(isOpen ? null : g.key)}
                className="w-full flex items-center justify-between py-2.5"
              >
                <span className="text-[var(--text-primary)] text-xs font-semibold capitalize">{monthLabel(g.key)}</span>
                <div className="flex items-center gap-2">
                  <span className="text-[var(--text-muted)] text-[11px] font-mono tabular-nums">{fmt(groupTotal)}</span>
                  <ChevronDown size={14} className={`text-[var(--text-muted)] transition-transform ${isOpen ? "rotate-180" : ""}`} />
                </div>
              </button>
              {isOpen && (
                <div className="pb-1">
                  {g.items.map((a) => (
                    <AdvanceHistoryRow key={a.id} a={a} t={t} onDelete={() => onDelete(a.id)} />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// YANGI: Admin uchun oylik "tabel" ko'rinishi — barcha ishchilar × kunlar jadvali,
// bitta ekranda kim qachon (to'liq/yarim/kelmagan/belgilanmagan) ekanini ko'rsatadi.
function MonthlyTimesheet({ employees, attendance }) {
  const { t, lang } = useApp();
  const localeTag = lang === "ru" ? "ru-RU" : lang === "en" ? "en-US" : "uz-UZ";
  const [monthOffset, setMonthOffset] = useState(0);

  const base = new Date();
  base.setDate(1);
  base.setMonth(base.getMonth() + monthOffset);
  const year = base.getFullYear();
  const month = base.getMonth();
  const monthLabel = base.toLocaleDateString(localeTag, { month: "long", year: "numeric" });
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthPrefix = `${year}-${String(month + 1).padStart(2, "0")}`;
  const days = Array.from({ length: daysInMonth }, (_, i) => i + 1);

  function cellInfo(emp, day) {
    const dateStr = `${monthPrefix}-${String(day).padStart(2, "0")}`;
    if (dateStr > todayISO()) return { kind: "future" };
    if (employeeJoinDate(emp) > dateStr) return { kind: "before-join" };
    const raw = attendance[emp.id]?.[dateStr];
    if (raw === undefined) return { kind: "unmarked" };
    return { kind: "marked", v: attEntryStatus(raw) };
  }

  function cellColor(info) {
    if (info.kind === "before-join" || info.kind === "future") return "transparent";
    if (info.kind === "unmarked") return "var(--bad-soft)";
    if (info.v === 1) return "var(--good)";
    if (info.v === 0.5) return "var(--warn)";
    return "var(--bad)";
  }

  function monthTotal(emp) {
    let sum = 0;
    for (const d of days) {
      const dateStr = `${monthPrefix}-${String(d).padStart(2, "0")}`;
      const raw = attendance[emp.id]?.[dateStr];
      if (raw !== undefined) sum += attEntryStatus(raw);
    }
    return sum;
  }

  return (
    <div className="card rounded-xl p-4">
      <div className="flex items-center justify-between mb-3">
        <button type="button" onClick={() => setMonthOffset((o) => o - 1)} className="w-8 h-8 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field">
          <ChevronLeft size={15} />
        </button>
        <span className="text-[var(--text-primary)] text-sm font-semibold capitalize">{monthLabel}</span>
        <button type="button" onClick={() => setMonthOffset((o) => Math.min(0, o + 1))} disabled={monthOffset === 0} className="w-8 h-8 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field disabled:opacity-30">
          <ChevronRight size={15} />
        </button>
      </div>

      {employees.length === 0 ? (
        <p className="text-[var(--text-muted)] text-sm text-center py-8">{t("noEmployees")}</p>
      ) : (
        <div data-noswipe="true" className="overflow-x-auto -mx-4 px-4">
          <table style={{ borderCollapse: "separate", borderSpacing: "2px" }}>
            <thead>
              <tr>
                <th className="sticky left-0 bg-[var(--bg-card)] text-left text-[10px] text-[var(--text-muted)] font-medium pr-2" style={{ minWidth: 108 }}>
                  {t("employee")}
                </th>
                {days.map((d) => (
                  <th key={d} className="text-[9px] text-[var(--text-faint)] font-medium" style={{ width: 20 }}>
                    {d}
                  </th>
                ))}
                <th className="text-[10px] text-[var(--text-muted)] font-medium pl-2" style={{ minWidth: 32 }}>Σ</th>
              </tr>
            </thead>
            <tbody>
              {employees.map((emp) => (
                <tr key={emp.id}>
                  <td className="sticky left-0 bg-[var(--bg-card)] text-[var(--text-primary)] text-xs pr-2 truncate" style={{ maxWidth: 108 }}>
                    {emp.name}
                  </td>
                  {days.map((d) => {
                    const info = cellInfo(emp, d);
                    return (
                      <td key={d} className="p-0">
                        <div className="rounded-[3px]" style={{ width: 18, height: 18, backgroundColor: cellColor(info) }} />
                      </td>
                    );
                  })}
                  <td className="text-[var(--text-secondary)] text-[11px] font-mono tabular-nums pl-2">{fmtDays(monthTotal(emp))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="flex items-center gap-3 mt-4 pt-3 border-t border-[var(--border-soft)] flex-wrap">
        <span className="flex items-center gap-1.5 text-[10px] text-[var(--text-muted)]">
          <span className="w-2.5 h-2.5 rounded-sm inline-block" style={{ backgroundColor: "var(--good)" }} /> {t("fullDay")}
        </span>
        <span className="flex items-center gap-1.5 text-[10px] text-[var(--text-muted)]">
          <span className="w-2.5 h-2.5 rounded-sm inline-block" style={{ backgroundColor: "var(--warn)" }} /> {t("halfDay")}
        </span>
        <span className="flex items-center gap-1.5 text-[10px] text-[var(--text-muted)]">
          <span className="w-2.5 h-2.5 rounded-sm inline-block" style={{ backgroundColor: "var(--bad)" }} /> {t("absent")}
        </span>
        <span className="flex items-center gap-1.5 text-[10px] text-[var(--text-muted)]">
          <span className="w-2.5 h-2.5 rounded-sm inline-block" style={{ backgroundColor: "var(--bad-soft)" }} /> {t("statusNone")}
        </span>
      </div>
    </div>
  );
}

// YANGI: Ba'zi qurilma/brauzerlarda tizimning umumiy ulashish oynasi (Web Share API)
// ishlamaydi yoki mavjud emas — bunday holatda kod jim-jim to'g'ridan-to'g'ri
// yuklab olishga o'tib ketardi, ulashish imkoniyati umuman ko'rinmasdi. Endi bu
// holatda o'zimizning kafolatlangan ulashish varag'imiz chiqadi: Telegram va
// WhatsApp'ga matnli xulosa bilan to'g'ridan-to'g'ri o'tish, Email orqali yuborish,
// yoki Excel faylini yuklab olish (keyin istalgan ilovaga qo'lda biriktirish mumkin).
const BRAND_ICON_SRC = {
  telegram: "data:image/png;base64,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",
  whatsapp: "data:image/png;base64,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",
  email: "data:image/png;base64,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",
  excel: "data:image/png;base64,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",
};

// YANGI: endi haqiqiy ilova belgilari (foydalanuvchi yuborgan rasmlar) ishlatiladi —
// fayl ichiga base64 qilib joylab qo'yildi, shuning uchun alohida rasm fayllarini
// loyihaga qo'shishning hojati yo'q, hammasi shu faylning o'zida ishlaydi.
function BrandIcon({ kind, size = 52 }) {
  const wrap = { width: size, height: size };
  if (kind === "email") {
    // Gmail belgisi faqat harf (orqa foni shaffof) — shuning uchun oq, yumaloq
    // burchakli "karta" fon beramiz, xuddi boshqa ilova belgilaridek ko'rinishi uchun.
    return (
      <div
        style={{ ...wrap, borderRadius: size * 0.27, background: "#fff", border: "1px solid #e3e5e8", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" }}
      >
        <img src={BRAND_ICON_SRC.email} alt="Email" style={{ width: "66%", height: "66%", objectFit: "contain" }} />
      </div>
    );
  }
  return (
    <img
      src={BRAND_ICON_SRC[kind]}
      alt={kind}
      style={{ ...wrap, borderRadius: kind === "excel" ? size * 0.22 : "50%", objectFit: "cover", display: "block" }}
    />
  );
}

// YANGI: Ba'zi qurilma/brauzerlarda tizimning umumiy ulashish oynasi (Web Share API)
// ishlamaydi yoki mavjud emas — bunday holatda o'zimizning ulashish varag'imiz chiqadi:
// Telegram va WhatsApp'ga matnli xulosa bilan o'tish, Email orqali yuborish,
// yoki Excel faylini yuklab olish.
function ShareMenu({ open, onClose, onTelegram, onWhatsapp, onEmail, onExcel }) {
  if (!open) return null;
  const items = [
    { kind: "telegram", label: "Telegram", onClick: onTelegram },
    { kind: "whatsapp", label: "WhatsApp", onClick: onWhatsapp },
    { kind: "email", label: "Gmail / Email", onClick: onEmail },
    { kind: "excel", label: "Excel", onClick: onExcel },
  ];
  return (
    <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-end" onClick={onClose}>
      <div
        className="w-full bg-[var(--bg-panel)] rounded-t-3xl p-5 pb-8"
        style={{ animation: "sheetSlideUp 0.25s ease-out" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="w-10 h-1 rounded-full bg-[var(--border-input)] mx-auto mb-4" />
        <div className="text-[var(--text-primary)] text-sm font-semibold mb-5 text-center">Ulashish</div>
        <div className="grid grid-cols-4 gap-3 max-w-sm mx-auto">
          {items.map((it) => (
            <button key={it.kind} type="button" onClick={it.onClick} className="flex flex-col items-center gap-2 active:scale-95 transition-transform">
              <BrandIcon kind={it.kind} />
              <span className="text-[10px] text-[var(--text-secondary)] text-center leading-tight">{it.label}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

function AdminApp({
  usersData, currentUser, onLogout, summaryFor,
  adminTab, setAdminTab,
  newEmp, setNewEmp, empError, addEmployee, deleteEmployee, updateEmployeeWage, resetEmployeePassword,
  attendance, attDate, setAttDate, markAttendance, bulkMarkAttendance,
  advances, advEmp, setAdvEmp, advForm, setAdvForm, addAdvance, deleteAdvance,
  changeOwnCredentials, updateAvatar, deleteOwnAccount, accent, setAccent, mode, setMode, fontScale, setFontScale, lang, setLang, enableNotifications,
  notifications, markAllNotificationsRead, markNotificationRead, linkTelegram,
  settleEmployee, undoLastSettlement, settlements, exportBackup,
}) {
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [showAddForm, setShowAddForm] = useState(false);
  const [addBusy, setAddBusy] = useState(false);
  const [notifOpen, setNotifOpen] = useState(false);
  const [empSearch, setEmpSearch] = useState("");
  const unreadCount = notifications.filter((n) => !n.is_read).length;
  const { t } = useApp();
  const myAdmin = usersData.admins[currentUser.username] || { avatar: null };
  const myEmployees = usersData.employees.filter((e) => e.owner === currentUser.username);
  const tabs = [
    { id: "employees", label: t("navEmployees"), icon: <Users size={18} /> },
    { id: "attendance", label: t("navAttendance"), icon: <Calendar size={18} /> },
    { id: "advances", label: t("navAdvances"), icon: <Wallet size={18} /> },
    { id: "report", label: t("navReport"), icon: <ClipboardList size={18} /> },
  ];
  const touchStartX = useRef(null);
  const touchStartY = useRef(null);
  function handleTouchStart(e) {
    // FIX: jadval kabi o'zining ichki gorizontal skrolli bo'lgan joylardan
    // boshlangan barmoq harakati endi tab almashtirish sifatida talqin qilinmaydi —
    // aks holda jadvalni o'ngga-chapga surish "boshqa bo'limga o'tib ketish"ga
    // olib kelardi. Bunday zonalar data-noswipe atributi bilan belgilanadi.
    if (e.target.closest && e.target.closest("[data-noswipe]")) {
      touchStartX.current = null;
      touchStartY.current = null;
      return;
    }
    touchStartX.current = e.touches[0].clientX;
    touchStartY.current = e.touches[0].clientY;
  }
  function handleTouchEnd(e) {
    if (touchStartX.current === null) return;
    const deltaX = e.changedTouches[0].clientX - touchStartX.current;
    const deltaY = e.changedTouches[0].clientY - touchStartY.current;
    touchStartX.current = null;
    touchStartY.current = null;
    if (Math.abs(deltaX) < 110) return;
    if (Math.abs(deltaX) < Math.abs(deltaY) * 1.5) return;
    const idx = tabs.findIndex((t) => t.id === adminTab);
    if (deltaX < 0 && idx < tabs.length - 1) setAdminTab(tabs[idx + 1].id);
    if (deltaX > 0 && idx > 0) setAdminTab(tabs[idx - 1].id);
  }

  const localeTag = lang === "ru" ? "ru-RU" : lang === "en" ? "en-US" : "uz-UZ";
  function buildReportRows() {
    return myEmployees.map((emp) => {
      const s = summaryFor(emp.id);
      return {
        [t("colEmployee")]: emp.name,
        [t("colDays")]: fmtDays(s.workedDays),
        [t("colCalculated")]: s.totalWage,
        [t("colAdvance")]: s.totalAdvance,
        [t("colRemaining")]: s.remaining,
      };
    });
  }
  function buildReportWorkbook() {
    const ws = XLSX.utils.json_to_sheet(buildReportRows());
    ws["!cols"] = [{ wch: 22 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, t("reportHeader").slice(0, 31));
    return wb;
  }
  function buildReportTextSummary() {
    const lines = myEmployees.map((emp) => {
      const s = summaryFor(emp.id);
      return `${emp.name}: ${fmtDays(s.workedDays)} kun · qoldiq ${fmt(s.remaining)}`;
    });
    return `${t("reportHeader")} — ${todayISO()}\n\n${lines.join("\n")}`;
  }
  // YANGI: "Ulashish" — avval brauzerning tabiiy ulashish oynasini (Web Share API)
  // ochishga harakat qiladi. Lekin ba'zi qurilma/brauzerlarda bu API umuman
  // mavjud emas yoki xato beradi — bunday holatda endi jim-jim yuklab olishga
  // O'TIB KETMAYDI, o'rniga o'zimizning kafolatlangan ulashish varag'i (ShareMenu)
  // ochiladi: Telegram/WhatsApp/Email/Excel variantlari bilan.
  async function shareReport() {
    if (typeof navigator.share === "function") {
      try {
        const wb = buildReportWorkbook();
        const fileName = `hisobot-${todayISO()}.xlsx`;
        const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const file = new File([blob], fileName, { type: blob.type });
        if (typeof navigator.canShare !== "function" || navigator.canShare({ files: [file] })) {
          await navigator.share({ files: [file], title: t("reportHeader") });
          return;
        }
        await navigator.share({ title: t("reportHeader"), text: buildReportTextSummary() });
        return;
      } catch (e) {
        if (e && e.name === "AbortError") return; // foydalanuvchi ulashishni bekor qildi
        // boshqa xato bo'lsa — pastdagi o'z menyumizga o'tamiz
      }
    }
    setShareMenuOpen(true);
  }

  function shareViaTelegram() {
    // t.me havolasi bu muhitda Telegram veb-saytini ochib yuborardi; tg:// esa
    // to'g'ridan-to'g'ri o'rnatilgan Telegram ilovasini ochadi.
    window.location.href = `tg://msg?text=${encodeURIComponent(buildReportTextSummary())}`;
    setShareMenuOpen(false);
  }
  function shareViaWhatsapp() {
    window.location.href = `https://wa.me/?text=${encodeURIComponent(buildReportTextSummary())}`;
    setShareMenuOpen(false);
  }
  function shareViaEmail() {
    window.location.href = `mailto:?subject=${encodeURIComponent(t("reportHeader"))}&body=${encodeURIComponent(buildReportTextSummary())}`;
    setShareMenuOpen(false);
  }
  function shareViaExcelDownload() {
    XLSX.writeFile(buildReportWorkbook(), `hisobot-${todayISO()}.xlsx`);
    setShareMenuOpen(false);
  }
  const [weekOffset, setWeekOffset] = useState(0);
  const [attView, setAttView] = useState("daily"); // "daily" | "monthly"
  const [shareMenuOpen, setShareMenuOpen] = useState(false);
  const [settleEmpId, setSettleEmpId] = useState(null);
  const [pendingBulk, setPendingBulk] = useState(null); // { status, label } yoki null
  const dateStrip = (() => {
    const today = new Date();
    const day = today.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    const monday = new Date(today);
    monday.setDate(today.getDate() + mondayOffset + weekOffset * 7);
    return Array.from({ length: 7 }, (_, i) => {
      const d = new Date(monday);
      d.setDate(monday.getDate() + i);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    });
  })();

  const bottomNav = (
    <nav className="fixed bottom-0 left-0 right-0 z-20 px-4" style={{ paddingBottom: "calc(env(safe-area-inset-bottom) + 16px)" }}>
      <div className="max-w-md mx-auto flex bg-[var(--bg-card)]/95 backdrop-blur-md border border-[var(--border)] shadow-[var(--shadow-card)] rounded-full px-1.5 py-1">
        {tabs.map((tab) => {
          const active = adminTab === tab.id;
          return (
            <button
              key={tab.id}
              type="button"
              onClick={() => setAdminTab(tab.id)}
              className="flex-1 flex justify-center"
            >
              <span
                className="w-full flex flex-col items-center gap-0.5 px-4 py-2.5 rounded-full text-[10px] font-medium transition-colors"
                style={active ? { backgroundColor: accent + "26", color: accent } : { color: "var(--text-muted)" }}
              >
                {tab.icon}
                {tab.label}
              </span>
            </button>
          );
        })}
      </div>
    </nav>
  );

  return (
    <>
      <ProfileDrawer
        open={drawerOpen}
        onClose={() => setDrawerOpen(false)}
        me={{ name: currentUser.username, username: currentUser.username, avatar: myAdmin.avatar }}
        roleLabel={t("adminPanel")}
        isAdmin={true}
        onDeleteAccount={deleteOwnAccount}
        onExportBackup={exportBackup}
        onLogout={onLogout}
        changeOwnCredentials={changeOwnCredentials}
        updateAvatar={updateAvatar}
        accent={accent} setAccent={setAccent}
        mode={mode} setMode={setMode}
        fontScale={fontScale} setFontScale={setFontScale}
        lang={lang} setLang={setLang}
        enableNotifications={enableNotifications}
        linkTelegram={linkTelegram}
      />
      <NotificationPanel
        open={notifOpen}
        onClose={() => setNotifOpen(false)}
        notifications={notifications}
        onMarkAllRead={() => markAllNotificationsRead()}
        onMarkRead={markNotificationRead}
      />
      {settleEmpId && (() => {
        const emp = myEmployees.find((e) => e.id === settleEmpId);
        if (!emp) return null;
        return (
          <SettleSheet
            emp={emp}
            previewFor={(d) => summaryFor(emp.id, d)}
            list={settlements[emp.id] || []}
            onSettle={(d, mode) => settleEmployee(emp.id, d, mode)}
            onUndo={() => undoLastSettlement(emp.id)}
            onClose={() => setSettleEmpId(null)}
          />
        );
      })()}
      <ShareMenu
        open={shareMenuOpen}
        onClose={() => setShareMenuOpen(false)}
        onTelegram={shareViaTelegram}
        onWhatsapp={shareViaWhatsapp}
        onEmail={shareViaEmail}
        onExcel={shareViaExcelDownload}
      />
      <Shell
        title={t("adminPanel")}
        userName={currentUser.username}
        avatar={myAdmin.avatar || null}
        onTitleClick={() => setDrawerOpen(true)}
        bottomNav={bottomNav}
        headerRight={
          <button
            type="button"
            onClick={() => setNotifOpen(true)}
            className="w-9 h-9 rounded-full flex items-center justify-center text-[var(--text-secondary)] hover:text-[var(--accent)] hover:bg-[var(--bg-card)] transition-colors shrink-0 relative"
            aria-label={t("notifications")}
          >
            <Bell size={18} />
            {unreadCount > 0 && (
              <span className="absolute top-1 right-1 min-w-[16px] h-4 px-1 rounded-full bg-[var(--bad)] text-white text-[9px] font-bold flex items-center justify-center">
                {unreadCount}
              </span>
            )}
          </button>
        }
      >
      <div key={adminTab} className="tab-transition" onTouchStart={handleTouchStart} onTouchEnd={handleTouchEnd}>
      {adminTab === "employees" && (
        <div className="space-y-5">
          {!showAddForm ? (
            <button
              type="button"
              onClick={() => setShowAddForm(true)}
              className="w-full flex items-center justify-center gap-1.5 py-3 rounded-xl border border-dashed border-[var(--border-input)] text-[var(--text-secondary)] text-sm font-medium hover:text-[var(--text-primary)] hover:border-[var(--accent)] transition-colors"
            >
              <UserPlus size={16} /> {t("addEmployeeHeader")}
            </button>
          ) : (
            <div className="card rounded-xl p-5">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-1.5 text-[var(--text-primary)] text-sm font-semibold">
                  <UserPlus size={15} /> {t("addEmployeeHeader")}
                </div>
                <button type="button" onClick={() => { setShowAddForm(false); }} className="text-[var(--text-muted)] hover:text-[var(--text-primary)] transition-colors">
                  <X size={16} />
                </button>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <Field label={t("fullName")} value={newEmp.name} onChange={(v) => setNewEmp({ ...newEmp, name: v })} />
                <MoneyField label={t("dailyWage")} value={newEmp.dailyWage} onChange={(v) => setNewEmp({ ...newEmp, dailyWage: v })} suffix="so'm" />
                <Field label={t("login")} value={newEmp.username} onChange={(v) => setNewEmp({ ...newEmp, username: v })} />
                <Field label={t("password")} type="password" value={newEmp.password} onChange={(v) => setNewEmp({ ...newEmp, password: v })} />
              </div>
              {empError && <p className="text-[var(--bad)] text-xs mt-3">{empError}</p>}
              <button
                type="button"
                disabled={addBusy}
                onClick={async () => { setAddBusy(true); const ok = await addEmployee(); setAddBusy(false); if (ok) setShowAddForm(false); }}
                className="mt-4 flex items-center gap-1.5 px-4 py-2 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity disabled:opacity-60"
                style={{ background: accentGradient(accent) }}
              >
                <Plus size={14} /> {addBusy ? t("loading") : t("add")}
              </button>
            </div>
          )}

          {myEmployees.length > 0 && (
            <div className="relative">
              <Search size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-[var(--text-muted)]" />
              <input
                type="text"
                value={empSearch}
                onChange={(e) => setEmpSearch(e.target.value)}
                placeholder="Ism bo'yicha qidirish..."
                className="w-full pl-9 pr-3 py-2.5 rounded-lg field text-[var(--text-primary)] text-sm outline-none focus:border-[var(--accent)] transition-colors"
              />
            </div>
          )}

          <div className="space-y-2">
            {(() => {
              const filteredEmployees = myEmployees.filter((emp) =>
                emp.name.toLowerCase().includes(empSearch.trim().toLowerCase())
              );
              if (myEmployees.length === 0) {
                return (
                  <EmptyState
                    icon={<Users size={20} />}
                    text={t("noEmployees")}
                    actionLabel={t("add")}
                    onAction={() => setShowAddForm(true)}
                  />
                );
              }
              if (filteredEmployees.length === 0) {
                return <p className="text-[var(--text-muted)] text-sm text-center py-8">Hech kim topilmadi</p>;
              }
              return filteredEmployees.map((emp) => (
                <EmployeeRow
                  key={emp.id}
                  emp={emp}
                  summary={summaryFor(emp.id)}
                  onDelete={() => deleteEmployee(emp.id)}
                  onUpdateWage={(w) => updateEmployeeWage(emp.id, w)}
                  onResetPassword={(pw) => resetEmployeePassword(emp.id, pw)}
                />
              ));
            })()}
          </div>
        </div>
      )}

      {adminTab === "attendance" && (
        <div className="space-y-4">
          <div className="flex gap-1.5 p-1 rounded-lg field">
            <button
              type="button"
              onClick={() => setAttView("daily")}
              className="flex-1 py-1.5 rounded-md text-xs font-medium transition-colors"
              style={attView === "daily" ? { backgroundColor: accent, color: "#12161c" } : { color: "var(--text-secondary)" }}
            >
              Kunlik
            </button>
            <button
              type="button"
              onClick={() => setAttView("monthly")}
              className="flex-1 py-1.5 rounded-md text-xs font-medium transition-colors"
              style={attView === "monthly" ? { backgroundColor: accent, color: "#12161c" } : { color: "var(--text-secondary)" }}
            >
              Oylik jadval
            </button>
          </div>

          {attView === "monthly" && (
            <MonthlyTimesheet employees={myEmployees} attendance={attendance} />
          )}

          {attView === "daily" && (
          <>
          <div className="card rounded-xl p-4">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5 text-[var(--text-primary)] text-sm font-semibold">
                <Calendar size={15} /> {t("markAttendanceHeader")}
              </div>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setWeekOffset((w) => w - 1)}
                className="shrink-0 w-7 h-14 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field transition-colors"
                aria-label="prev week"
              >
                <ChevronLeft size={16} />
              </button>
              <div className="flex-1 grid grid-cols-7 gap-1.5">
                {dateStrip.map((d) => {
                  const dt = new Date(d + "T00:00:00");
                  const weekday = dt.toLocaleDateString(localeTag, { weekday: "short" });
                  const isToday = d === todayISO();
                  const isSelected = d === attDate;
                  const isFuture = d > todayISO();
                  const dayEmployees = myEmployees.filter((emp) => employeeJoinDate(emp) <= d);
                  const markedCount = dayEmployees.filter((emp) => attendance[emp.id]?.[d] !== undefined).length;
                  const hasData = dayEmployees.length > 0 && !isFuture;
                  const dotColor = !hasData
                    ? "transparent"
                    : markedCount === 0
                      ? "var(--bad)"
                      : markedCount < dayEmployees.length
                        ? "var(--warn)"
                        : "var(--good)";
                  const style = isFuture
                    ? { backgroundColor: "var(--bg-app)", color: "var(--text-faint)", border: "1px solid var(--border-input)", opacity: 0.45 }
                    : isToday
                      ? { backgroundColor: "var(--good)", color: "#0e1712", boxShadow: isSelected ? `0 0 0 2px var(--bg-card), 0 0 0 4px ${accent}` : "none" }
                      : isSelected
                        ? { backgroundColor: accent, color: "#12161c" }
                        : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" };
                  return (
                    <button
                      key={d}
                      type="button"
                      disabled={isFuture}
                      onClick={() => setAttDate(d)}
                      className="flex flex-col items-center justify-center py-2 rounded-lg text-xs transition-colors disabled:cursor-not-allowed"
                      style={style}
                    >
                      <span className="text-[10px] opacity-80 capitalize">{weekday}</span>
                      <span className="text-sm font-semibold">{dt.getDate()}</span>
                      <span
                        className="w-2 h-2 rounded-full mt-1"
                        style={{
                          backgroundColor: dotColor,
                          border: hasData && markedCount === 0 ? "none" : "1px solid rgba(0,0,0,0.08)",
                        }}
                      />
                    </button>
                  );
                })}
              </div>
              <button
                type="button"
                onClick={() => setWeekOffset((w) => w + 1)}
                className="shrink-0 w-7 h-14 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field transition-colors"
                aria-label="next week"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          </div>

          {(() => {
            const visibleEmployees = myEmployees.filter((emp) => employeeJoinDate(emp) <= attDate);
            // hisob-kitob qilingan (qulflangan) davrdagi ishchilarni ommaviy belgilashdan chiqaramiz
            const editableEmployees = visibleEmployees.filter((emp) => {
              const c = summaryFor(emp.id)?.cutoff;
              return !(c && attDate <= c);
            });
            return (
              <>
          {editableEmployees.length > 0 && attDate <= todayISO() && (
            <>
              <div className="flex gap-1.5">
                <button
                  type="button"
                  onClick={() => setPendingBulk({ status: 1, label: t("markAllFull"), icon: <CheckCircle2 size={13} />, tone: "good" })}
                  className="flex-1 flex items-center justify-center gap-1 py-2 rounded-lg text-[11px] font-medium border border-[var(--good)]/40 text-[var(--good)] bg-transparent hover:bg-[var(--good-soft)] transition-colors"
                >
                  <CheckCircle2 size={13} /> {t("markAllFull")}
                </button>
                <button
                  type="button"
                  onClick={() => setPendingBulk({ status: 0.5, label: t("markAllHalf"), icon: <Calendar size={13} />, tone: "warn" })}
                  className="flex-1 flex items-center justify-center gap-1 py-2 rounded-lg text-[11px] font-medium border border-[var(--warn)]/40 text-[var(--warn)] bg-transparent hover:bg-[var(--warn-soft)] transition-colors"
                >
                  <Calendar size={13} /> {t("markAllHalf")}
                </button>
                <button
                  type="button"
                  onClick={() => setPendingBulk({ status: 0, label: t("markAllAbsent"), icon: <XCircle size={13} />, tone: "bad" })}
                  className="flex-1 flex items-center justify-center gap-1 py-2 rounded-lg text-[11px] font-medium border border-[var(--bad)]/40 text-[var(--bad)] bg-transparent hover:bg-[var(--bad-soft)] transition-colors"
                >
                  <XCircle size={13} /> {t("markAllAbsent")}
                </button>
              </div>

              {pendingBulk && (
                <div
                  className="flex items-center justify-between gap-3 rounded-lg p-3 mt-1.5"
                  style={{
                    backgroundColor: `var(--${pendingBulk.tone}-soft)`,
                    border: `1px solid var(--${pendingBulk.tone})`,
                  }}
                >
                  <span className="text-xs font-medium flex items-center gap-1.5" style={{ color: `var(--${pendingBulk.tone})` }}>
                    {pendingBulk.icon} {editableEmployees.length} ta ishchiga "{pendingBulk.label}" qo'llansinmi?
                  </span>
                  <div className="flex gap-1.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => { bulkMarkAttendance(editableEmployees.map((e) => e.id), pendingBulk.status); setPendingBulk(null); }}
                      className="px-3 py-1.5 rounded-md text-white text-[11px] font-semibold"
                      style={{ backgroundColor: `var(--${pendingBulk.tone})` }}
                    >
                      {t("yesConfirm")}
                    </button>
                    <button
                      type="button"
                      onClick={() => setPendingBulk(null)}
                      className="px-3 py-1.5 rounded-md field text-[var(--text-secondary)] text-[11px] font-medium"
                    >
                      {t("cancel")}
                    </button>
                  </div>
                </div>
              )}
            </>
          )}

          <div className="space-y-2">
            {visibleEmployees.length === 0 && (
              <EmptyState
                icon={<Users size={20} />}
                text={t("noEmployees")}
                actionLabel={t("add")}
                onAction={() => { setAdminTab("employees"); setShowAddForm(true); }}
              />
            )}
            {visibleEmployees.length > 0 && attDate > todayISO() && (
              <p className="text-[var(--warn)] text-xs text-center py-2 card rounded-lg">{t("futureDateWarning")}</p>
            )}
            {visibleEmployees.map((emp) => {
              const hasEntry = attendance[emp.id]?.[attDate] !== undefined;
              const st = hasEntry ? attEntryStatus(attendance[emp.id]?.[attDate]) : null;
              const isFuture = attDate > todayISO();
              const lockedUntil = summaryFor(emp.id)?.cutoff;
              const locked = !!(lockedUntil && attDate <= lockedUntil);

              function nextStatus() {
                if (st === null) return 1;
                if (st === 1) return 0.5;
                if (st === 0.5) return 0;
                return 1;
              }

              return (
                <AttendanceStatusRow
                  key={emp.id}
                  emp={emp}
                  status={st}
                  isFuture={isFuture}
                  locked={locked}
                  onCycle={() => markAttendance(emp.id, nextStatus())}
                />
              );
            })}
          </div>
              </>
            );
          })()}
          </>
          )}
        </div>
      )}

      {adminTab === "advances" && (
        <div className="space-y-5">
          <div className="card rounded-xl p-5">
            <div className="flex items-center gap-1.5 text-[var(--text-primary)] text-sm font-semibold mb-4">
              <Wallet size={15} /> {t("giveAdvanceHeader")}
            </div>
            <div className="grid grid-cols-2 gap-1.5 mb-3">
              <button
                type="button"
                onClick={() => setAdvForm({ ...advForm, type: "avans" })}
                className="py-2 rounded-lg text-xs font-medium transition-colors"
                style={(advForm.type || "avans") === "avans" ? { backgroundColor: accent, color: "#12161c" } : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }}
              >
                {t("typeAvans")}
              </button>
              <button
                type="button"
                onClick={() => setAdvForm({ ...advForm, type: "salary" })}
                className="py-2 rounded-lg text-xs font-medium transition-colors"
                style={advForm.type === "salary" ? { backgroundColor: accent, color: "#12161c" } : { backgroundColor: "var(--bg-app)", color: "var(--text-secondary)", border: "1px solid var(--border-input)" }}
              >
                {t("typeSalary")}
              </button>
            </div>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div className="col-span-2">
                <label className="block text-xs text-[var(--text-secondary)] mb-1.5">{t("employee")}</label>
                <select value={advEmp} onChange={(e) => setAdvEmp(e.target.value)}
                  className="w-full px-3 py-2.5 rounded-lg field text-[var(--text-primary)] text-sm outline-none focus:border-[var(--accent)]">
                  <option value="">{t("selectPlaceholder")}</option>
                  {myEmployees.map((e) => <option key={e.id} value={e.id}>{e.name}</option>)}
                </select>
              </div>
              <MoneyField label={t("amount")} value={advForm.amount} onChange={(v) => setAdvForm({ ...advForm, amount: v })} suffix="so'm" />
              <Field label={t("date")} type="date" value={advForm.date} onChange={(v) => setAdvForm({ ...advForm, date: v })} />
              <div className="col-span-2">
                <Field label={t("note")} value={advForm.note} onChange={(v) => setAdvForm({ ...advForm, note: v })} />
              </div>
            </div>
            <button type="button" onClick={addAdvance} className="flex items-center gap-1.5 px-4 py-2 rounded-lg text-white text-xs font-semibold hover:opacity-90 transition-opacity" style={{ background: accentGradient(accent) }}>
              <Plus size={14} /> {(advForm.type === "salary") ? t("addSalaryPayment") : t("addAdvance")}
            </button>
          </div>

          {advEmp && (
            <AdvanceHistoryCard
              list={advances[advEmp] || []}
              onDelete={(advId) => deleteAdvance(advEmp, advId)}
            />
          )}
        </div>
      )}

      {adminTab === "report" && (
        <div className="space-y-4">
          {(() => {
            const totalOwed = myEmployees.reduce((sum, emp) => {
              const r = summaryFor(emp.id).remaining;
              return sum + (r > 0 ? r : 0);
            }, 0);
            return (
              <div className="card rounded-xl p-6 text-center">
                <div className="text-[var(--text-muted)] text-xs mb-1.5">Jami to'lash kerak</div>
                <AnimatedAmount
                  value={totalOwed}
                  formatter={fmt}
                  className={`text-3xl font-bold font-mono tabular-nums ${totalOwed > 0 ? "text-[var(--bad)]" : "text-[var(--good)]"}`}
                />
              </div>
            );
          })()}
        <div className="card rounded-xl overflow-hidden">
          <div className="p-5 pb-3 flex items-center justify-between gap-2">
            <div className="text-[var(--text-primary)] text-sm font-semibold flex items-center gap-1.5 min-w-0">
              <ClipboardList size={15} className="shrink-0" /> <span className="truncate">{t("reportHeader")}</span>
            </div>
            {myEmployees.length > 0 && (
              <button
                type="button"
                onClick={shareReport}
                aria-label="Ulashish"
                title="Ulashish"
                className="w-9 h-9 rounded-lg flex items-center justify-center text-white hover:opacity-90 transition-opacity shrink-0"
                style={{ background: accentGradient(accent) }}
              >
                <Share2 size={16} />
              </button>
            )}
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-[var(--text-muted)] text-xs border-t border-[var(--border)]">
                  <th className="text-left font-medium py-2.5 px-5">{t("colEmployee")}</th>
                  <th className="text-right font-medium py-2.5 px-3">{t("colDays")}</th>
                  <th className="text-right font-medium py-2.5 px-3">{t("colCalculated")}</th>
                  <th className="text-right font-medium py-2.5 px-3">{t("colAdvance")}</th>
                  <th className="text-right font-medium py-2.5 px-3">{t("colRemaining")}</th>
                  <th className="text-right font-medium py-2.5 px-5">Qarz</th>
                </tr>
              </thead>
              <tbody>
                {myEmployees.map((emp) => {
                  const s = summaryFor(emp.id);
                  return (
                    <tr key={emp.id} className="border-t border-[var(--border-soft)]">
                      <td className="py-2.5 px-5 text-[var(--text-primary)]">
                        <div>{emp.name}</div>
                        <button
                          type="button"
                          onClick={() => setSettleEmpId(emp.id)}
                          className="text-[11px] font-semibold mt-0.5"
                          style={{ color: accent }}
                        >
                          Hisoblash
                        </button>
                      </td>
                      <td className="py-2.5 px-3 text-right text-[var(--text-secondary)] font-mono tabular-nums">{fmtDays(s.workedDays)}</td>
                      <td className="py-2.5 px-3 text-right text-[var(--text-secondary)] font-mono tabular-nums">{fmt(s.totalWage)}</td>
                      <td className="py-2.5 px-3 text-right text-[var(--bad)] font-mono tabular-nums">-{fmt(s.totalAdvance)}</td>
                      <td className="py-2.5 px-3 text-right font-semibold font-mono tabular-nums">
                        {s.remaining === 0 ? (
                          <Check size={14} className="inline text-[var(--warn)]" />
                        ) : s.remaining < 0 ? (
                          <span className="text-[var(--text-faint)]">—</span>
                        ) : (
                          <span className="text-[var(--good)]">{fmt(s.remaining)}</span>
                        )}
                      </td>
                      <td className="py-2.5 px-5 text-right font-semibold font-mono tabular-nums">
                        {s.remaining < 0 ? (
                          <span className="text-[var(--bad)]">{fmt(Math.abs(s.remaining))}</span>
                        ) : (
                          <span className="text-[var(--text-faint)]">—</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
                  {myEmployees.length === 0 && (
                  <tr><td colSpan={6} className="py-8 text-center text-[var(--text-muted)]">{t("noData")}</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
        </div>
      )}
      </div>
      </Shell>
    </>
  );
}

function EmployeeApp({
  currentUser, usersData, summaryFor, onLogout,
  changeOwnCredentials, updateAvatar, deleteOwnAccount, accent, setAccent, mode, setMode, fontScale, setFontScale, lang, setLang, linkTelegram, enableNotifications,
  notifications, markAllNotificationsRead, markNotificationRead,
}) {
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [notifOpen, setNotifOpen] = useState(false);
  const unreadCount = notifications.filter((n) => !n.is_read).length;
  const [empTab, setEmpTab] = useState("umumiy");
  const [attMonthOffset, setAttMonthOffset] = useState(0);
  const [advMonthKey, setAdvMonthKey] = useState(todayISO().slice(0, 7));
  const advScrollRef = useRef(null);
  const advMonthRefs = useRef({});
  const advScrollRafRef = useRef(null);

  function handleAdvScroll() {
    if (advScrollRafRef.current) return;
    advScrollRafRef.current = requestAnimationFrame(() => {
      advScrollRafRef.current = null;
      const container = advScrollRef.current;
      if (!container) return;
      const containerCenter = container.getBoundingClientRect().left + container.clientWidth / 2;
      let closestKey = null;
      let closestDist = Infinity;
      Object.entries(advMonthRefs.current).forEach(([key, el]) => {
        if (!el) return;
        const rect = el.getBoundingClientRect();
        const center = rect.left + rect.width / 2;
        const dist = Math.abs(center - containerCenter);
        if (dist < closestDist) { closestDist = dist; closestKey = key; }
      });
      if (closestKey && closestKey !== advMonthKey) setAdvMonthKey(closestKey);
    });
  }

  function scrollAdvToMonth(key) {
    const el = advMonthRefs.current[key];
    if (el) el.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
  }

  useEffect(() => {
    if (empTab !== "avanslar") return;
    let raf2 = null;
    const raf1 = requestAnimationFrame(() => {
      raf2 = requestAnimationFrame(() => {
        const el = advMonthRefs.current[advMonthKey];
        if (el) el.scrollIntoView({ behavior: "auto", inline: "center", block: "nearest" });
      });
    });
    return () => {
      cancelAnimationFrame(raf1);
      if (raf2) cancelAnimationFrame(raf2);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [empTab]);
  const { t } = useApp();
  const s = summaryFor(currentUser.id);
  const attDays = Object.entries(s.att)
    .map(([date, raw]) => [date, attEntryStatus(raw), attEntryWage(raw, s.emp, date)])
    .sort((a, b) => (a[0] < b[0] ? 1 : -1));
  const empTabs = [
    { id: "umumiy", label: "Umumiy", icon: <Home size={18} /> },
    { id: "davomat", label: "Davomat", icon: <Calendar size={18} /> },
    { id: "avanslar", label: "Avanslar", icon: <Wallet size={18} /> },
  ];
  const empBottomNav = (
    <nav className="fixed bottom-0 left-0 right-0 z-20 px-4" style={{ paddingBottom: "calc(env(safe-area-inset-bottom) + 16px)" }}>
      <div className="max-w-md mx-auto flex bg-[var(--bg-card)]/95 backdrop-blur-md border border-[var(--border)] shadow-[var(--shadow-card)] rounded-full px-1.5 py-1">
        {empTabs.map((tab) => {
          const active = empTab === tab.id;
          return (
            <button
              key={tab.id}
              type="button"
              onClick={() => setEmpTab(tab.id)}
              className="flex-1 flex justify-center"
            >
              <span
                className="w-full flex flex-col items-center gap-0.5 px-4 py-2.5 rounded-full text-[10px] font-medium transition-colors"
                style={active ? { backgroundColor: accent + "26", color: accent } : { color: "var(--text-muted)" }}
              >
                {tab.icon}
                {tab.label}
              </span>
            </button>
          );
        })}
      </div>
    </nav>
  );

  return (
    <>
      <ProfileDrawer
        open={drawerOpen}
        onClose={() => setDrawerOpen(false)}
        me={{ name: s.emp.name, username: s.emp.username, avatar: s.emp.avatar }}
        roleLabel={t("employeePanel")}
        isAdmin={false}
        onDeleteAccount={deleteOwnAccount}
        onLogout={onLogout}
        changeOwnCredentials={changeOwnCredentials}
        updateAvatar={updateAvatar}
        accent={accent} setAccent={setAccent}
        mode={mode} setMode={setMode}
        fontScale={fontScale} setFontScale={setFontScale}
        lang={lang} setLang={setLang}
        linkTelegram={linkTelegram}
        enableNotifications={enableNotifications}
      />
      <NotificationPanel
        open={notifOpen}
        onClose={() => setNotifOpen(false)}
        notifications={notifications}
        onMarkAllRead={() => markAllNotificationsRead()}
        onMarkRead={markNotificationRead}
      />
      <Shell
        title={t("employeePanel")}
        userName={currentUser.name}
        avatar={s.emp.avatar || null}
        onTitleClick={() => setDrawerOpen(true)}
        bottomNav={empBottomNav}
        headerRight={
          <button
            type="button"
            onClick={() => setNotifOpen(true)}
            className="w-9 h-9 rounded-full flex items-center justify-center text-[var(--text-secondary)] hover:text-[var(--accent)] hover:bg-[var(--bg-card)] transition-colors shrink-0 relative"
            aria-label={t("notifications")}
          >
            <Bell size={18} />
            {unreadCount > 0 && (
              <span className="absolute top-1 right-1 min-w-[16px] h-4 px-1 rounded-full bg-[var(--bad)] text-white text-[9px] font-bold flex items-center justify-center">
                {unreadCount}
              </span>
            )}
          </button>
        }
      >
        {empTab === "umumiy" && (
          <div className="tab-transition space-y-3">
            <div className="card rounded-xl p-6 text-center">
              <div className="flex items-center justify-center gap-1.5 text-[var(--text-muted)] text-xs mb-1.5">
                <Wallet size={13} /> {t("statRemainingSalary")}
              </div>
              <AnimatedAmount value={s.remaining} formatter={fmt} className="text-3xl font-bold font-mono tabular-nums text-[var(--good)]" />
              {s.cutoff && (
                <div className="text-[var(--text-muted)] text-[11px] mt-2">
                  Oxirgi hisob: {s.cutoff}{s.opening !== 0 ? ` · oldingi qoldiq ${fmt(s.opening)}` : ""}
                </div>
              )}
            </div>
            <div className="card rounded-xl overflow-hidden grid grid-cols-2">
              <StatCell label={t("statWorkedDays")} value={fmtDays(s.workedDays)} icon={<Calendar size={12} />} border="r b" />
              <StatCell label={t("statDailyWage")} value={fmt(s.emp.dailyWage)} icon={<Wallet size={12} />} border="b" />
              <StatCell label={t("typeAvans")} value={fmt(s.totalAvans)} tone="bad" icon={<TrendingDown size={12} />} border="r" />
              <StatCell label={t("typeSalary")} value={fmt(s.totalSalaryPaid)} tone="bad" icon={<Wallet size={12} />} />
            </div>
          </div>
        )}
        {empTab === "davomat" && (() => {
          const localeTag = lang === "ru" ? "ru-RU" : lang === "en" ? "en-US" : "uz-UZ";
          const base = new Date();
          base.setDate(1);
          base.setMonth(base.getMonth() + attMonthOffset);
          const year = base.getFullYear();
          const month = base.getMonth();
          const monthLabel = base.toLocaleDateString(localeTag, { month: "long", year: "numeric" });
          const daysInMonth = new Date(year, month + 1, 0).getDate();
          const startWeekday = (new Date(year, month, 1).getDay() + 6) % 7;
          const attMap = Object.fromEntries(attDays.map(([date, v, wage]) => [date, [v, wage]]));
          const cells = [];
          for (let i = 0; i < startWeekday; i++) cells.push(null);
          for (let d = 1; d <= daysInMonth; d++) cells.push(d);
          const monthPrefix = `${year}-${String(month + 1).padStart(2, "0")}`;
          const monthEntries = attDays.filter(([date]) => date.startsWith(monthPrefix));
          const monthWorkedDays = monthEntries.reduce((sum, [, v]) => sum + v, 0);
          const monthWage = monthEntries.reduce((sum, [, v, w]) => sum + v * w, 0);
          const weekdayLabels = localeTag === "uz-UZ" ? ["D", "S", "CH", "P", "J", "SH", "Y"] : ["M", "T", "W", "T", "F", "S", "S"];

          return (
            <div className="tab-transition space-y-3">
              <div className="card rounded-xl p-4">
                <div className="flex items-center justify-between mb-3">
                  <button type="button" onClick={() => setAttMonthOffset((o) => o - 1)} className="w-8 h-8 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field">
                    <ChevronLeft size={15} />
                  </button>
                  <span className="text-[var(--text-primary)] text-sm font-semibold capitalize">{monthLabel}</span>
                  <button type="button" onClick={() => setAttMonthOffset((o) => Math.min(0, o + 1))} disabled={attMonthOffset === 0} className="w-8 h-8 rounded-lg flex items-center justify-center text-[var(--text-muted)] hover:text-[var(--text-primary)] field disabled:opacity-30">
                    <ChevronRight size={15} />
                  </button>
                </div>
                <div className="grid grid-cols-7 gap-1 mb-1.5">
                  {weekdayLabels.map((w, i) => (
                    <div key={i} className="text-center text-[9px] text-[var(--text-muted)] uppercase">{w}</div>
                  ))}
                </div>
                <div className="grid grid-cols-7 gap-1">
                  {cells.map((d, i) => {
                    if (d === null) return <div key={i} />;
                    const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
                    const entry = attMap[dateStr];
                    const isFuture = dateStr > todayISO();
                    let style = { backgroundColor: "var(--bg-app)", color: "var(--text-faint)", border: "1px solid var(--border-input)" };
                    if (!isFuture && entry) {
                      const v = entry[0];
                      if (v === 1) style = { backgroundColor: "var(--good)", color: "#0e1712" };
                      else if (v === 0.5) style = { backgroundColor: "var(--warn)", color: "#1a1608" };
                      else style = { backgroundColor: "var(--bad)", color: "#1c0e0c" };
                    }
                    return (
                      <div key={i} className="aspect-square rounded-md flex items-center justify-center text-[11px] font-medium" style={style}>
                        {d}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="card rounded-xl overflow-hidden grid grid-cols-2">
                <StatCell label="Shu oy ishlagan" value={fmtDays(monthWorkedDays)} icon={<Calendar size={12} />} border="r" />
                <StatCell label="Shu oy hisoblangan" value={fmt(monthWage)} tone="good" icon={<Wallet size={12} />} />
              </div>
            </div>
          );
        })()}
        {empTab === "avanslar" && (() => {
          const localeTag = lang === "ru" ? "ru-RU" : lang === "en" ? "en-US" : "uz-UZ";
          const monthNamesUz = ["yanvar", "fevral", "mart", "aprel", "may", "iyun", "iyul", "avgust", "sentabr", "oktabr", "noyabr", "dekabr"];

          const PAST_MONTHS = 24;
          const FUTURE_MONTHS = 1;
          const months = Array.from({ length: PAST_MONTHS + FUTURE_MONTHS + 1 }, (_, i) => {
            const d = new Date();
            d.setDate(1);
            d.setMonth(d.getMonth() - PAST_MONTHS + i);
            const mIdx = d.getMonth();
            const yy = d.getFullYear();
            const name = localeTag === "uz-UZ" ? monthNamesUz[mIdx] : d.toLocaleDateString(localeTag, { month: "long" });
            const label = (mIdx === 11 || mIdx === 0) ? `${yy} / ${name}` : name;
            return { key: `${yy}-${String(mIdx + 1).padStart(2, "0")}`, label };
          });

          const filteredAdv = s.advList.filter((a) => a.date.startsWith(advMonthKey));

          return (
            <div className="space-y-3 tab-transition">
              <div
                ref={advScrollRef}
                onScroll={handleAdvScroll}
                className="flex gap-2 overflow-x-auto pb-1 -mx-5 px-5"
                style={{ scrollbarWidth: "none", scrollSnapType: "x mandatory" }}
              >
                {months.map((m) => {
                  const active = advMonthKey === m.key;
                  return (
                    <button
                      key={m.key}
                      ref={(el) => { advMonthRefs.current[m.key] = el; }}
                      type="button"
                      onClick={() => scrollAdvToMonth(m.key)}
                      className="flex items-center justify-center text-center capitalize rounded-lg"
                      style={{
                        scrollSnapAlign: "center",
                        flex: "0 0 calc((100% - 16px) / 3)",
                        height: active ? 52 : 40,
                        border: active ? `2px solid ${accent}` : "1px solid var(--border-input)",
                        color: active ? "var(--text-primary)" : "var(--text-faint)",
                        fontWeight: active ? 700 : 500,
                        fontSize: active ? 16 : 12,
                        backgroundColor: "var(--bg-card)",
                        opacity: active ? 1 : 0.45,
                        filter: active ? "none" : "blur(1.2px)",
                        transform: active ? "scale(1)" : "scale(0.86)",
                        transition: "opacity 0.25s ease, filter 0.25s ease, transform 0.25s ease, height 0.25s ease, font-size 0.25s ease, border-color 0.25s ease",
                      }}
                    >
                      {m.label}
                    </button>
                  );
                })}
              </div>

              {filteredAdv.length === 0 && (
                <div className="card rounded-xl p-8 flex flex-col items-center gap-2 text-center">
                  <Wallet size={20} className="text-[var(--text-faint)]" />
                  <p className="text-[var(--text-muted)] text-xs">{t("noAdvancesYet")}</p>
                </div>
              )}
              {filteredAdv.slice().reverse().map((a) => (
                <div key={a.id} className="card rounded-xl p-2.5 flex items-center gap-2.5">
                  <div
                    className="w-8 h-8 rounded-full flex items-center justify-center shrink-0"
                    style={a.type === "salary" ? { backgroundColor: "var(--good-soft)", color: "var(--good)" } : { backgroundColor: "var(--warn-soft)", color: "var(--warn)" }}
                  >
                    {a.type === "salary" ? <Wallet size={14} /> : <TrendingDown size={14} />}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="text-[var(--text-primary)] text-sm font-semibold font-mono tabular-nums mb-1">{fmt(a.amount)}</div>
                    <span
                      className="text-[9px] font-semibold uppercase tracking-wide px-1.5 py-0.5 rounded-full inline-block mr-1.5"
                      style={a.type === "salary" ? { backgroundColor: "var(--good-soft)", color: "var(--good)" } : { backgroundColor: "var(--warn-soft)", color: "var(--warn)" }}
                    >
                      {a.type === "salary" ? t("typeSalary") : t("typeAvans")}
                    </span>
                    <span className="text-[var(--text-muted)] text-xs">{a.date}</span>
                  </div>
                  {a.note && (
                    <div className="text-[var(--text-primary)] text-sm font-bold text-right shrink-0 max-w-[38%] truncate">{a.note}</div>
                  )}
                </div>
              ))}
            </div>
          );
        })()}
      </Shell>
    </>
  );
}

export default function WorkforceApp() {
  return (
    <AppErrorBoundary>
      <WorkforceAppInner />
    </AppErrorBoundary>
  );
}

function WorkforceAppInner() {
  const [loading, setLoading] = useState(true);
  const [usersData, setUsersData] = useState(null);
  const [attendance, setAttendance] = useState({});
  const [advances, setAdvances] = useState({});
  const [settlements, setSettlements] = useState({}); // { [employeeId]: [hisob-kitoblar, sana bo'yicha o'sish tartibida] }
  const [session, setSession] = useState(null);
  const [currentUser, setCurrentUserState] = useState(null);
  const [telegramPromptOpen, setTelegramPromptOpen] = useState(false);
  const [telegramPromptPhase, setTelegramPromptPhase] = useState("prompt"); // prompt | waiting | success
  const [telegramLinkBusy, setTelegramLinkBusy] = useState(false);
  const telegramPollRef = useRef(null);
  const [isOnline, setIsOnline] = useState(() => (typeof navigator !== "undefined" ? navigator.onLine : true));
  const [retrying, setRetrying] = useState(false);

  useEffect(() => {
    function goOnline() { setIsOnline(true); }
    function goOffline() { setIsOnline(false); }
    window.addEventListener("online", goOnline);
    window.addEventListener("offline", goOffline);
    return () => {
      window.removeEventListener("online", goOnline);
      window.removeEventListener("offline", goOffline);
    };
  }, []);

  async function retryConnection() {
    setRetrying(true);
    setIsOnline(navigator.onLine);
    if (navigator.onLine && session) {
      try { await loadAllData(session.user); } catch (e) {}
    }
    setRetrying(false);
  }

  function setCurrentUser(user) {
    setCurrentUserState(user);
  }

  // YANGI: "Meni eslab qol" endi haqiqatan ishlaydi — belgilangan bo'lsa,
  // loginni (parolni EMAS) brauzerda saqlab qo'yamiz va keyingi safar
  // kirish ekrani avtomatik to'ldirib beradi.
  const [rememberMe, setRememberMe] = useState(() => {
    try { return localStorage.getItem("nazorat_remember_me") !== "0"; } catch (e) { return true; }
  });
  const [loginForm, setLoginForm] = useState(() => {
    try {
      const savedUsername = localStorage.getItem("nazorat_remembered_username") || "";
      return { username: savedUsername, password: "" };
    } catch (e) {
      return { username: "", password: "" };
    }
  });
  const [loginError, setLoginError] = useState("");
  const [loginErrorTick, setLoginErrorTick] = useState(0);
  const [lockSeconds, setLockSeconds] = useState(0);
  const [lockTick, setLockTick] = useState(0);
  const [loginBusy, setLoginBusy] = useState(false);

  const [adminTab, setAdminTab] = useState("employees");
  const [newEmp, setNewEmp] = useState({ name: "", username: "", password: "", dailyWage: "" });
  const [empError, setEmpError] = useState("");
  const [attDate, setAttDate] = useState(todayISO());
  const [advEmp, setAdvEmp] = useState("");
  const [advForm, setAdvForm] = useState({ amount: "", date: todayISO(), note: "", type: "avans" });
  const [notifications, setNotifications] = useState([]);
  const [accent, setAccent] = useState(ACCENT_PRESETS[2].value);
  const [mode, setMode] = useState(() => {
    try {
      return localStorage.getItem("app-mode") || "dark";
    } catch (e) {
      return "dark";
    }
  });
  const [fontScale, setFontScale] = useState(100);
  const [lang, setLang] = useState("uz");

  // ============================================================================
  // MA'LUMOTLARNI YUKLASH — endi app_storage o'rniga real Supabase Auth +
  // profiles/attendance/advances jadvallaridan o'qiladi (RLS himoyasi bilan).
  // ============================================================================

  async function loadAllData(sessionUser) {
    // 1) O'z profilimni olamiz (rolimni aniqlash uchun)
    const { data: myProfile, error: myErr } = await supabase
      .from("profiles")
      .select("*")
      .eq("id", sessionUser.id)
      .single();
    if (myErr || !myProfile) {
      await supabase.auth.signOut();
      setCurrentUserState(null);
      setLoading(false);
      return;
    }

    if (myProfile.role === "admin") {
      const { data: employeesRaw } = await supabase
        .from("profiles")
        .select("*")
        .eq("owner_id", sessionUser.id);
      const employees = (employeesRaw || []).map((e) => ({
        id: e.id, name: e.display_name || e.username, username: e.username,
        dailyWage: Number(e.daily_wage || 0), avatar: e.avatar_url,
        owner: myProfile.username, wageHistory: e.wage_history || [],
      }));
      const empIds = employees.map((e) => e.id);

      const attMap = {};
      const advMap = {};
      const setMap = {};
      if (empIds.length > 0) {
        const { data: attRows } = await supabase.from("attendance").select("*").in("employee_id", empIds);
        (attRows || []).forEach((r) => {
          if (!attMap[r.employee_id]) attMap[r.employee_id] = {};
          attMap[r.employee_id][r.date] = { v: Number(r.status), wage: Number(r.wage_at_time) };
        });
        const { data: advRows } = await supabase.from("advances").select("*").in("employee_id", empIds).order("created_at", { ascending: true });
        (advRows || []).forEach((r) => {
          if (!advMap[r.employee_id]) advMap[r.employee_id] = [];
          advMap[r.employee_id].push({ id: r.id, amount: Number(r.amount), date: r.date, note: r.note, type: r.type });
        });
        const { data: setRows } = await supabase.from("settlements").select("*").in("employee_id", empIds).order("closed_through", { ascending: true });
        (setRows || []).forEach((r) => {
          if (!setMap[r.employee_id]) setMap[r.employee_id] = [];
          setMap[r.employee_id].push(mapSettlementRow(r));
        });
      }

      setUsersData({
        admins: { [myProfile.username]: { avatar: myProfile.avatar_url } },
        employees,
      });
      setAttendance(attMap);
      setAdvances(advMap);
      setSettlements(setMap);
      setCurrentUserState({ role: "admin", name: makeT(lang)("admin"), username: myProfile.username, id: myProfile.id });
      await loadNotifications(myProfile.id);
    } else {
      // Ishchi: o'z ma'lumotlarini va admin (owner)ining avatarini olamiz
      const { data: ownerProfile } = await supabase.from("profiles").select("username").eq("id", myProfile.owner_id).maybeSingle();
      const emp = {
        id: myProfile.id, name: myProfile.display_name || myProfile.username, username: myProfile.username,
        dailyWage: Number(myProfile.daily_wage || 0), avatar: myProfile.avatar_url,
        owner: ownerProfile?.username || "", wageHistory: myProfile.wage_history || [],
      };
      const { data: attRows } = await supabase.from("attendance").select("*").eq("employee_id", myProfile.id);
      const attMap = {};
      (attRows || []).forEach((r) => { attMap[r.date] = { v: Number(r.status), wage: Number(r.wage_at_time) }; });
      const { data: advRows } = await supabase.from("advances").select("*").eq("employee_id", myProfile.id).order("created_at", { ascending: true });
      const advList = (advRows || []).map((r) => ({ id: r.id, amount: Number(r.amount), date: r.date, note: r.note, type: r.type }));

      setUsersData({ admins: {}, employees: [emp] });
      setAttendance({ [myProfile.id]: attMap });
      setAdvances({ [myProfile.id]: advList });
      const { data: mySetRows } = await supabase.from("settlements").select("*").eq("employee_id", myProfile.id).order("closed_through", { ascending: true });
      setSettlements({ [myProfile.id]: (mySetRows || []).map(mapSettlementRow) });
      await loadNotifications(myProfile.id);
      setCurrentUserState({ role: "employee", id: myProfile.id, name: emp.name, owner: emp.owner });
    }
    setLoading(false);
  }

  useEffect(() => {
    let mounted = true;
    supabase.auth.getSession().then(({ data }) => {
      if (!mounted) return;
      setSession(data.session);
      if (data.session) loadAllData(data.session.user);
      else setLoading(false);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => { mounted = false; sub.subscription.unsubscribe(); };
  }, []);

  useEffect(() => {
    try { localStorage.setItem("app-mode", mode); } catch (e) {}
  }, [mode]);

  useEffect(() => {
    if (typeof document !== "undefined" && document.documentElement) {
      document.documentElement.style.fontSize = fontScale + "%";
    }
    return () => {
      if (typeof document !== "undefined" && document.documentElement) {
        document.documentElement.style.fontSize = "";
      }
    };
  }, [fontScale]);

  // FIX: avval bu yerda "attendance/advances/profiles" jadvallarini kuzatib
  // turadigan realtime subscription bor edi, u har o'zgarishda BUTUN
  // ma'lumotni qayta yuklardi. Muammo shu edi: har bir amal (masalan davomat
  // belgilash) ham o'zi to'g'ridan-to'g'ri qayta yuklardi, HAM shu realtime
  // signal orqali yana bir marta yuklanardi — ikkitasi bir vaqtda to'qnashib,
  // tez-tez bosilganda ilova "qotib qolar" edi. Endi faqat aniq amaldan keyingi
  // bitta yuklash yetarli, realtime esa olib tashlandi.

  useEffect(() => {
    if (!currentUser) return;
    const channel = supabase
      .channel(`notif_${currentUser.id}`)
      .on(
        "postgres_changes",
        { event: "INSERT", schema: "public", table: "notifications", filter: `admin_id=eq.${currentUser.id}` },
        (payload) => setNotifications((prev) => [payload.new, ...prev])
      )
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [currentUser?.id]);

  async function loadNotifications(adminId) {
    const { data } = await supabase
      .from("notifications")
      .select("*")
      .eq("admin_id", adminId)
      .order("created_at", { ascending: false })
      .limit(30);
    setNotifications(data || []);
  }

  async function markAllNotificationsRead() {
    if (!currentUser) return;
    await supabase.from("notifications").update({ is_read: true }).eq("admin_id", currentUser.id).eq("is_read", false);
    setNotifications((prev) => prev.map((n) => ({ ...n, is_read: true })));
  }

  async function markNotificationRead(id) {
    if (!currentUser) return;
    setNotifications((prev) => prev.map((n) => (n.id === id ? { ...n, is_read: true } : n)));
    await supabase.from("notifications").update({ is_read: true }).eq("id", id);
  }

  // ============================================================================
  // LOGIN / RO'YXATDAN O'TISH
  // ============================================================================

  async function handleLogin() {
    setLoginError("");
    const username = loginForm.username.trim();
    const password = loginForm.password;
    if (!username || !password) {
      setLoginError(makeT(lang)("wrongLogin"));
      setLoginErrorTick((t) => t + 1);
      return;
    }
    setLoginBusy(true);
    try {
      // YANGI: urinishdan oldin shu login vaqtincha bloklanmaganini tekshiramiz.
      const { data: lockData } = await supabase.rpc("check_login_lock", { p_username: username });
      const lock = Array.isArray(lockData) ? lockData[0] : lockData;
      if (lock && lock.locked) {
        setLockSeconds(lock.seconds_left);
        setLockTick((x) => x + 1);
        setLoginBusy(false);
        return;
      }
      const { data: email, error: lookupErr } = await supabase.rpc("email_for_username", { p_username: username });
      if (lookupErr || !email) {
        const { data: failData } = await supabase.rpc("record_login_failure", { p_username: username });
        const fail = Array.isArray(failData) ? failData[0] : failData;
        if (fail && fail.locked) { setLockSeconds(fail.seconds_left); setLockTick((x) => x + 1); }
        else { setLoginError(makeT(lang)("wrongLogin")); setLoginErrorTick((t) => t + 1); }
        setLoginBusy(false);
        return;
      }
      const { data: signInData, error: signInErr } = await supabase.auth.signInWithPassword({ email, password });
      if (signInErr || !signInData?.session) {
        const { data: failData } = await supabase.rpc("record_login_failure", { p_username: username });
        const fail = Array.isArray(failData) ? failData[0] : failData;
        if (fail && fail.locked) { setLockSeconds(fail.seconds_left); setLockTick((x) => x + 1); }
        else { setLoginError(makeT(lang)("wrongLogin")); setLoginErrorTick((t) => t + 1); }
        setLoginBusy(false);
        return;
      }
      await supabase.rpc("record_login_success", { p_username: username });
      try {
        if (rememberMe) {
          localStorage.setItem("nazorat_remembered_username", username);
          localStorage.setItem("nazorat_remember_me", "1");
        } else {
          localStorage.removeItem("nazorat_remembered_username");
          localStorage.setItem("nazorat_remember_me", "0");
        }
      } catch (e) {}
      setSession(signInData.session);
      await loadAllData(signInData.session.user);
    } catch (err) {
      setLoginError(String(err && err.message ? err.message : err));
    }
    setLoginBusy(false);
  }

  async function registerAdmin(newUsername, newPassword) {
    try {
      const { error } = await invokeFn("register-admin", { username: newUsername, password: newPassword });
      if (error) {
        return { error };
      }
      const { data: signInData, error: signInErr } = await supabase.auth.signInWithPassword({
        email: `${newUsername.toLowerCase()}@nazoratplus.internal`,
        password: newPassword,
      });
      if (signInErr || !signInData?.session) {
        return { error: "Ro'yxatdan o'tildi, lekin avtomatik kirishda xato. Iltimos, qo'lda kiring." };
      }
      setSession(signInData.session);
      await loadAllData(signInData.session.user);
      setTelegramPromptOpen(true);
      return {};
    } catch (err) {
      return { error: String(err && err.message ? err.message : err) };
    }
  }

  function logout() {
    supabase.auth.signOut();
    setCurrentUserState(null);
    setUsersData(null);
    setSession(null);
    setLoginForm({ username: "", password: "" });
  }

  // ============================================================================
  // HISOB-KITOB (o'zgarmagan)
  // ============================================================================

  // YANGI: oxirgi hisob-kitobdan (settlement) keyingi joriy davr bo'yicha hisoblaydi.
  // throughDate berilsa — shu sanagacha (shu sana ham kiradi) bo'lgan davr hisoblanadi
  // (hisob-kitob qilishdan oldin natijani ko'rsatish uchun). Oldingi davrdan qolgan
  // qoldiq (carry_over) boshlang'ich balans sifatida qo'shiladi.
  function summaryFor(empId, throughDate = null) {
    const emp = usersData.employees.find((x) => x.id === empId);
    if (!emp) return null;
    const setList = settlements[empId] || [];
    const last = setList.length ? setList[setList.length - 1] : null;
    const cutoff = last ? last.closedThrough : null;
    const opening = last ? Number(last.carryOver) : 0;
    const inPeriod = (date) => (!cutoff || date > cutoff) && (!throughDate || date <= throughDate);

    const att = attendance[empId] || {};
    let workedDays = 0;
    let totalWage = 0;
    for (const [date, raw] of Object.entries(att)) {
      if (!inPeriod(date)) continue;
      const v = attEntryStatus(raw);
      workedDays += v;
      totalWage += v * attEntryWage(raw, emp, date);
    }
    const advList = advances[empId] || [];
    const periodAdv = advList.filter((a) => inPeriod(a.date));
    const totalAvans = periodAdv.filter((a) => a.type !== "salary").reduce((sum, a) => sum + Number(a.amount), 0);
    const totalSalaryPaid = periodAdv.filter((a) => a.type === "salary").reduce((sum, a) => sum + Number(a.amount), 0);
    const totalAdvance = totalAvans + totalSalaryPaid;
    return { emp, workedDays, totalWage, totalAdvance, totalAvans, totalSalaryPaid, opening, cutoff, remaining: opening + totalWage - totalAdvance, advList, att };
  }

  function cutoffFor(empId) {
    const list = settlements[empId] || [];
    return list.length ? list[list.length - 1].closedThrough : null;
  }

  // mode: "pay"  — qoldiq miqdorida "ish haqi to'lovi" yozadi va hisobni yopadi (qoldiq 0);
  //       "carry" — hisobni yopadi, qoldiq/qarz keyingi hisobga o'tadi.
  async function settleEmployee(empId, closedThrough, mode) {
    if (!closedThrough) return { error: "Sanani tanlang" };
    if (closedThrough > todayISO()) return { error: "Kelajak sanasi uchun hisoblab bo'lmaydi" };
    const sm = summaryFor(empId, closedThrough);
    if (!sm) return { error: "Ishchi topilmadi" };
    if (sm.cutoff && closedThrough <= sm.cutoff) {
      return { error: `Oxirgi hisob ${sm.cutoff} sanasida qilingan. Undan keyingi sanani tanlang.` };
    }
    let paid = sm.totalAdvance;
    let remaining = sm.remaining;
    let newAdvance = null;
    if (mode === "pay" && sm.remaining > 0) {
      const { data: adv, error: advErr } = await supabase.from("advances").insert({
        employee_id: empId, amount: sm.remaining, date: closedThrough, note: "Oylik hisoblash", type: "salary",
      }).select().single();
      if (advErr) return { error: advErr.message };
      newAdvance = adv;
      paid += sm.remaining;
      remaining = 0;
    }
    const { data, error } = await supabase.from("settlements").insert({
      employee_id: empId,
      closed_through: closedThrough,
      opening_balance: sm.opening,
      worked_days: sm.workedDays,
      total_wage: sm.totalWage,
      total_paid: paid,
      remaining,
      carry_over: remaining,
    }).select().single();
    if (error) {
      // hisob yopilmadi — yozilgan to'lovni ham qaytarib olamiz, ma'lumot chala qolmasin
      if (newAdvance) await supabase.from("advances").delete().eq("id", newAdvance.id);
      return { error: error.message };
    }
    if (newAdvance) {
      setAdvances((prev) => ({
        ...prev,
        [empId]: [...(prev[empId] || []), { id: newAdvance.id, amount: Number(newAdvance.amount), date: newAdvance.date, note: newAdvance.note, type: newAdvance.type }],
      }));
    }
    setSettlements((prev) => ({ ...prev, [empId]: [...(prev[empId] || []), mapSettlementRow(data)] }));
    return {};
  }

  // Faqat eng oxirgi hisob-kitobni bekor qilish mumkin (xato qilingan bo'lsa).
  async function undoLastSettlement(empId) {
    const list = settlements[empId] || [];
    const last = list[list.length - 1];
    if (!last) return {};
    const { error } = await supabase.from("settlements").delete().eq("id", last.id);
    if (error) return { error: error.message };
    setSettlements((prev) => ({ ...prev, [empId]: (prev[empId] || []).filter((x) => x.id !== last.id) }));
    return {};
  }

  // ============================================================================
  // ISHCHILARNI BOSHQARISH — endi Edge Function + RPC orqali
  // ============================================================================

  async function addEmployee() {
    setEmpError("");
    if (!newEmp.name || !newEmp.username || !newEmp.password || !newEmp.dailyWage) {
      setEmpError(makeT(lang)("fillAllFields"));
      return false;
    }
    const { error } = await invokeFn("create-employee", {
      name: newEmp.name, username: newEmp.username, password: newEmp.password, dailyWage: Number(newEmp.dailyWage),
    });
    if (error) {
      setEmpError(error);
      return false;
    }
    setNewEmp({ name: "", username: "", password: "", dailyWage: "" });
    confetti({ particleCount: 100, spread: 70, origin: { y: 0.6 } });
    await loadAllData(session.user);
    return true;
  }

  async function deleteEmployee(id) {
    const { error } = await invokeFn("delete-employee", { employeeId: id });
    if (error) { console.error(error); return; }
    if (advEmp === id) setAdvEmp("");
    await loadAllData(session.user);
  }

  async function updateEmployeeWage(id, newWage) {
    const { error } = await supabase.rpc("admin_update_employee_wage", { p_employee_id: id, p_new_wage: newWage });
    if (error) { console.error(error); return; }
    await loadAllData(session.user);
  }

  async function resetEmployeePassword(id, newPassword) {
    const { error } = await invokeFn("reset-employee-password", { employeeId: id, newPassword });
    if (error) return { error };
    return {};
  }

  // ============================================================================
  // DAVOMAT
  // ============================================================================

  async function markAttendance(empId, status) {
    if (attDate > todayISO()) return;
    const lockedUntil = cutoffFor(empId);
    if (lockedUntil && attDate <= lockedUntil) return; // hisob-kitob qilingan davr qulflangan
    const currentRaw = attendance[empId]?.[attDate];
    const currentStatus = currentRaw !== undefined ? attEntryStatus(currentRaw) : null;
    // FIX: avval har bosishda BUTUN ma'lumot serverdan qayta so'ralardi —
    // tez-tez bosilganda so'rovlar to'planib, ilova "qotib qolardi". Endi
    // ekrandagi holatni DARHOL (optimistik) yangilaymiz, DB yozuvi orqa fonda
    // ketadi. Xato bo'lsa, eski holatga qaytariladi.
    const prevAttendance = attendance;
    if (currentStatus === status) {
      setAttendance((prev) => {
        const dayMap = { ...(prev[empId] || {}) };
        delete dayMap[attDate];
        return { ...prev, [empId]: dayMap };
      });
      const { error } = await supabase.from("attendance").delete().eq("employee_id", empId).eq("date", attDate);
      if (error) setAttendance(prevAttendance);
    } else {
      const emp = usersData.employees.find((e) => e.id === empId);
      const wage = Number(emp ? emp.dailyWage : 0);
      setAttendance((prev) => ({
        ...prev,
        [empId]: { ...(prev[empId] || {}), [attDate]: { v: status, wage } },
      }));
      const { error } = await supabase.from("attendance").upsert(
        { employee_id: empId, date: attDate, status, wage_at_time: wage },
        { onConflict: "employee_id,date" }
      );
      if (error) setAttendance(prevAttendance);
    }
  }

  async function bulkMarkAttendance(empIds, status) {
    if (attDate > todayISO()) return;
    empIds = empIds.filter((id) => { const c = cutoffFor(id); return !(c && attDate <= c); });
    if (empIds.length === 0) return;
    const rows = empIds.map((id) => {
      const emp = usersData.employees.find((e) => e.id === id);
      const wage = emp ? wageForDate(emp, attDate) : 0;
      return { employee_id: id, date: attDate, status, wage_at_time: Number(wage || 0) };
    });
    const prevAttendance = attendance;
    setAttendance((prev) => {
      const next = { ...prev };
      rows.forEach((r) => {
        next[r.employee_id] = { ...(next[r.employee_id] || {}), [r.date]: { v: r.status, wage: r.wage_at_time } };
      });
      return next;
    });
    const { error } = await supabase.from("attendance").upsert(rows, { onConflict: "employee_id,date" });
    if (error) setAttendance(prevAttendance);
  }

  // ============================================================================
  // AVANS / TO'LOV
  // ============================================================================

  async function addAdvance() {
    const amountNum = Number(advForm.amount);
    if (!advEmp || !advForm.amount || !Number.isFinite(amountNum) || amountNum <= 0) return;
    const lockedUntil = cutoffFor(advEmp);
    if (lockedUntil && advForm.date <= lockedUntil) {
      alert(`Bu sana hisoblangan davrga kiradi (${lockedUntil} gacha). Undan keyingi sanani tanlang.`);
      return;
    }
    const { data, error } = await supabase.from("advances").insert({
      employee_id: advEmp, amount: amountNum, date: advForm.date, note: advForm.note || null, type: advForm.type || "avans",
    }).select().single();
    setAdvForm({ amount: "", date: todayISO(), note: "", type: advForm.type || "avans" });
    if (!error && data) {
      setAdvances((prev) => ({
        ...prev,
        [advEmp]: [...(prev[advEmp] || []), { id: data.id, amount: Number(data.amount), date: data.date, note: data.note, type: data.type }],
      }));
    }
  }

  async function deleteAdvance(empId, advId) {
    setAdvances((prev) => ({ ...prev, [empId]: (prev[empId] || []).filter((a) => a.id !== advId) }));
    await supabase.from("advances").delete().eq("id", advId);
  }

  // YANGI: barcha biznes ma'lumotini (ishchilar, davomat, avanslar, hisob-kitoblar)
  // bitta JSON faylga saqlab, qurilmaga yuklab beradi. Hech narsa serverga yuborilmaydi —
  // faqat allaqachon ilovada yuklangan ma'lumot faylga aylantiriladi.
  function exportBackup() {
    const backup = {
      exportedAt: new Date().toISOString(),
      admin: currentUser?.username || null,
      employees: usersData?.employees || [],
      attendance,
      advances,
      settlements,
    };
    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `nazorat-zaxira-${todayISO()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // ============================================================================
  // PROFIL / XAVFSIZLIK
  // ============================================================================

  async function verifyCurrentPassword(password) {
    try {
      const { data: userRes } = await supabase.auth.getUser();
      const email = userRes?.user?.email;
      if (!email) return false;
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      return !error;
    } catch (e) {
      return false;
    }
  }

  async function changeOwnCredentials(newUsername, newPasswordOrNull, currentPassword) {
    const ok = await verifyCurrentPassword(currentPassword);
    if (!ok) return { error: makeT(lang)("errWrongCurrentPassword") };

    if (newPasswordOrNull) {
      const { error: pwErr } = await supabase.auth.updateUser({ password: newPasswordOrNull });
      if (pwErr) return { error: pwErr.message };
    }
    if (currentUser.role === "admin") {
      const taken = usersData.employees.some((e) => e.username === newUsername) ;
      // eslint-disable-next-line no-unused-vars
    }
    const { error: rpcErr } = await supabase.rpc("update_own_profile", { new_username: newUsername, new_avatar_url: null });
    if (rpcErr) return { error: rpcErr.message.includes("duplicate") ? makeT(lang)("errLoginTaken") : rpcErr.message };

    await loadAllData(session.user);
    return {};
  }

  async function deleteOwnAccount(currentPassword) {
    const ok = await verifyCurrentPassword(currentPassword);
    if (!ok) return { error: makeT(lang)("errWrongCurrentPassword") };
    const { error } = await invokeFn("delete-account");
    if (error) return { error };
    logout();
    return {};
  }

  async function updateAvatar(dataUrl) {
    const { error } = await supabase.rpc("update_own_profile", { new_username: null, new_avatar_url: dataUrl });
    if (error) { console.error(error); return; }
    await loadAllData(session.user);
  }

  // ============================================================================
  // BILDIRISHNOMALAR (push)
  // ============================================================================

  async function linkTelegram() {
    try {
      const { data, error } = await invokeFn("telegram-link-start");
      if (error) {
        return { error };
      }
      window.open(data.linkUrl, "_blank");
      startTelegramLinkPolling();
      return {};
    } catch (e) {
      return { error: String(e?.message || e) };
    }
  }

  // "Start" tugmasi Telegram tomonda bosilishini kuzatib boradi (har 2 soniyada
  // profilni tekshiradi). Bosilgach — tick + "Xush kelibsiz" ko'rsatiladi.
  function startTelegramLinkPolling() {
    setTelegramPromptPhase("waiting");
    let attempts = 0;
    if (telegramPollRef.current) clearInterval(telegramPollRef.current);
    telegramPollRef.current = setInterval(async () => {
      attempts++;
      try {
        const { data: userRes } = await supabase.auth.getUser();
        const uid = userRes?.user?.id;
        if (uid) {
          const { data: prof } = await supabase.from("profiles").select("telegram_chat_id").eq("id", uid).maybeSingle();
          if (prof?.telegram_chat_id) {
            clearInterval(telegramPollRef.current);
            telegramPollRef.current = null;
            setTelegramPromptPhase("success");
            setTimeout(() => { setTelegramPromptOpen(false); setTelegramPromptPhase("prompt"); }, 1800);
            return;
          }
        }
      } catch (e) {}
      if (attempts >= 60) { // ~2 daqiqa
        clearInterval(telegramPollRef.current);
        telegramPollRef.current = null;
        setTelegramPromptPhase("prompt");
      }
    }, 2000);
  }

  useEffect(() => {
    return () => { if (telegramPollRef.current) clearInterval(telegramPollRef.current); };
  }, []);

  async function enableNotifications() {
    if (!currentUser) return;
    if (!("serviceWorker" in navigator) || !("PushManager" in window)) {
      alert("Bu qurilma/brauzer bildirishnomani qo'llab-quvvatlamaydi");
      return;
    }
    try {
      const permission = await Notification.requestPermission();
      if (permission !== "granted") return;
      const reg = await navigator.serviceWorker.ready;
      let sub = await reg.pushManager.getSubscription();
      if (!sub) {
        sub = await reg.pushManager.subscribe({
          userVisibleOnly: true,
          applicationServerKey: urlBase64ToUint8Array(VAPID_PUBLIC_KEY),
        });
      }
      const subJson = sub.toJSON();
      await supabase.from("push_subscriptions").upsert(
        { admin_id: currentUser.id, subscription: subJson, endpoint: subJson.endpoint },
        { onConflict: "endpoint" }
      );
      alert("Bildirishnoma yoqildi!");
    } catch (e) {
      console.error(e);
    }
  }

  const t = makeT(lang);

  let screen;
  if (loading) {
    screen = (
      <div className="min-h-screen flex flex-col items-center justify-center gap-4 bg-[var(--bg-app)]">
        <img src="/logo.svg" alt={t("appName")} className="w-16 h-16 animate-pulse" />
        <div className="text-[var(--text-primary)] font-semibold text-base tracking-tight">{t("appName")}</div>
        <div className="text-[var(--text-muted)] text-xs">{t("loading")}</div>
      </div>
    );
  } else if (!currentUser || !usersData) {
    screen = (
      <LoginScreen
        loginForm={loginForm}
        setLoginForm={setLoginForm}
        loginError={loginError}
        loginBusy={loginBusy}
        onSubmit={handleLogin}
        onRegister={registerAdmin}
        rememberMe={rememberMe}
        setRememberMe={setRememberMe}
        lockSeconds={lockSeconds}
        lockTick={lockTick}
        loginErrorTick={loginErrorTick}
      />
    );
  } else if (currentUser.role === "admin") {
    screen = (
      <AdminApp
        usersData={usersData}
        currentUser={currentUser}
        onLogout={logout}
        summaryFor={summaryFor}
        adminTab={adminTab}
        setAdminTab={setAdminTab}
        newEmp={newEmp}
        setNewEmp={setNewEmp}
        empError={empError}
        addEmployee={addEmployee}
        deleteEmployee={deleteEmployee}
        updateEmployeeWage={updateEmployeeWage}
        resetEmployeePassword={resetEmployeePassword}
        attendance={attendance}
        attDate={attDate}
        setAttDate={setAttDate}
        markAttendance={markAttendance}
        bulkMarkAttendance={bulkMarkAttendance}
        advances={advances}
        advEmp={advEmp}
        setAdvEmp={setAdvEmp}
        advForm={advForm}
        setAdvForm={setAdvForm}
        addAdvance={addAdvance}
        deleteAdvance={deleteAdvance}
        changeOwnCredentials={changeOwnCredentials}
        updateAvatar={updateAvatar}
        deleteOwnAccount={deleteOwnAccount}
        accent={accent} setAccent={setAccent}
        mode={mode} setMode={setMode}
        fontScale={fontScale} setFontScale={setFontScale}
        lang={lang} setLang={setLang}
        enableNotifications={enableNotifications}
        notifications={notifications}
        markAllNotificationsRead={markAllNotificationsRead}
        markNotificationRead={markNotificationRead}
        linkTelegram={linkTelegram}
        settleEmployee={settleEmployee}
        undoLastSettlement={undoLastSettlement}
        settlements={settlements}
        exportBackup={exportBackup}
      />
    );
  } else {
    screen = (
      <EmployeeApp
        currentUser={currentUser}
        usersData={usersData}
        summaryFor={summaryFor}
        onLogout={logout}
        changeOwnCredentials={changeOwnCredentials}
        updateAvatar={updateAvatar}
        deleteOwnAccount={deleteOwnAccount}
        accent={accent} setAccent={setAccent}
        mode={mode} setMode={setMode}
        fontScale={fontScale} setFontScale={setFontScale}
        lang={lang} setLang={setLang}
        linkTelegram={linkTelegram}
        enableNotifications={enableNotifications}
        notifications={notifications}
        markAllNotificationsRead={markAllNotificationsRead}
        markNotificationRead={markNotificationRead}
      />
    );
  }

  return (
    <AppContext.Provider value={{ accent, lang, t }}>
      {!isOnline && (
        <div className="sticky top-0 z-[70] bg-[var(--bad)] text-white text-xs font-medium flex items-center justify-center gap-3 py-2 px-4 flex-wrap text-center">
          <span>Internet aloqasi yo'q</span>
          <button type="button" onClick={retryConnection} disabled={retrying} className="underline font-semibold disabled:opacity-60">
            {retrying ? "Tekshirilmoqda..." : "Qayta urinish"}
          </button>
        </div>
      )}
      <div
        className="font-sans"
        style={{
          ...PALETTES[mode],
          "--accent": accent,
          "--good": mode === "dark" ? "#4fb587" : "#2f9463",
          "--good-soft": "rgba(79,181,135,0.15)",
          "--bad": mode === "dark" ? "#e2685f" : "#d1453b",
          "--bad-soft": mode === "dark" ? "rgba(226,104,95,0.14)" : "rgba(209,69,59,0.1)",
          "--warn": mode === "dark" ? "#d9a53c" : "#b8811f",
          "--warn-soft": "rgba(217,165,60,0.15)",
          fontFamily: "'Manrope', system-ui, -apple-system, sans-serif",
        }}
      >
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700&family=Manrope:ital@0;1&subset=cyrillic&display=swap');
          .font-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
          @keyframes fadeSlideIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
          }
          .tab-transition { animation: fadeSlideIn 0.28s ease-out; }

          /* YANGI: davomat holatini bosganda badge biroz kattalashib qaytadi —
             rang almashishi "qattiq" tuyulmasligi uchun. */
          @keyframes statusPop {
            0% { transform: scale(1); }
            45% { transform: scale(1.14); }
            100% { transform: scale(1); }
          }
          .status-pop { animation: statusPop 0.28s ease; }

          /* YANGI: bildirishnoma pastdan chiqadigan varaq (bottom sheet)
             ochilganda pastdan tepaga sirg'alib chiqadi. */
          @keyframes sheetSlideUp {
            from { transform: translateY(100%); }
            to { transform: translateY(0); }
          }

          /* YANGI: login xato bo'lganda input maydoni titraydi (vibratsiyaga o'xshab). */
          @keyframes shake {
            10%, 90% { transform: translateX(-1px); }
            20%, 80% { transform: translateX(2px); }
            30%, 50%, 70% { transform: translateX(-4px); }
            40%, 60% { transform: translateX(4px); }
          }
          .shake-anim { animation: shake 0.4s ease; }

          /* YANGI: tekis (flat), zamonaviy uslub. Har bir "kartochka" nozik
             border va bitta yumshoq soya bilan ajralib turadi — ikki tomonlama
             neumorphic soyalar olib tashlandi, chunki ular qora fonda
             notekis va xira ko'rinar edi. */
          .card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            box-shadow: var(--shadow-card);
            transition: border-color 0.15s ease, transform 0.15s ease;
          }
          .field {
            background: var(--bg-app);
            border: 1px solid var(--border-input);
          }
          button.card:active {
            transform: scale(0.98);
          }

          button:focus-visible,
          select:focus-visible,
          a:focus-visible,
          [tabindex]:focus-visible {
            outline: 2px solid var(--accent);
            outline-offset: 2px;
            border-radius: 6px;
          }
        `}</style>
        {screen}
        {telegramPromptOpen && (
          <TelegramPromptModal
            phase={telegramPromptPhase}
            busy={telegramLinkBusy}
            onLink={async () => {
              setTelegramLinkBusy(true);
              await linkTelegram();
              setTelegramLinkBusy(false);
            }}
            onSkip={() => {
              if (telegramPollRef.current) { clearInterval(telegramPollRef.current); telegramPollRef.current = null; }
              setTelegramPromptOpen(false);
              setTelegramPromptPhase("prompt");
            }}
          />
        )}
      </div>
    </AppContext.Provider>
  );
}
