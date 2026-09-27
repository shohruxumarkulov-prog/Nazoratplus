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
function IconInput({ icon, type = "text", value, onChange, placeholder, onKeyDown, autoFocus, showToggle, toggleIcon, onToggle }) {
  const [focused, setFocused] = useState(false);
  const baseShadow = "inset -6px -6px 10px rgba(255,255,255,0.95), inset 6px 6px 10px rgba(184,190,204,0.45)";
  const focusShadow = "inset 3px 3px 6px rgba(20,60,140,0.35), inset -3px -3px 6px rgba(70,130,220,0.25)";
  return (
    <div className="relative">
      <span
        className="absolute left-4 top-1/2 -translate-y-1/2 pointer-events-none transition-colors duration-300"
        style={{ color: focused ? "#1a56b0" : "#909090" }}
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
        onFocus={() => setFocused(true)}
        onBlur={() => setFocused(false)}
        className={`w-full pl-11 ${showToggle ? "pr-11" : "pr-4"} py-3.5 rounded-2xl bg-[#e8e8e8] text-[#4a4a4a] text-sm font-medium outline-none border-none transition-all duration-300 placeholder:text-[#a3a3a3]`}
        style={{ boxShadow: focused ? focusShadow : baseShadow, transform: focused ? "translateY(-2px)" : "translateY(0)" }}
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

function LoginScreen({ loginForm, setLoginForm, loginError, onSubmit, onRegister, loginBusy, rememberMe, setRememberMe }) {
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
  const { t } = useApp();

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
                onChange={(v) => setLoginForm({ ...loginForm, username: v })}
                onKeyDown={(e) => { if (e.key === "Enter") onSubmit(); }}
                placeholder={t("login")}
                autoFocus={!registering}
              />
              <IconInput
                icon={<Lock size={17} />}
                type={showPassword ? "text" : "password"}
                value={loginForm.password}
                onChange={(v) => setLoginForm({ ...loginForm, password: v })}
                onKeyDown={(e) => { if (e.key === "Enter") onSubmit(); }}
                placeholder={t("password")}
                showToggle
                toggleIcon={showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                onToggle={() => setShowPassword((v) => !v)}
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

            {loginError && <p className="text-xs mt-3 text-center" style={{ color: "#a10f0f" }}>{loginError}</p>}

            <button
              type="button"
              disabled={loginBusy}
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
              {loginBusy ? t("loading") : t("loginBtn")}
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
function AttendanceStatusRow({ emp, status, isFuture, onCycle }) {
  const { t } = useApp();
  const [pulse, setPulse] = useState(false);
  const pulseTimer = useRef(null);

  useEffect(() => () => { if (pulseTimer.current) clearTimeout(pulseTimer.current); }, []);

  function handleClick() {
    if (isFuture) return;
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
      disabled={isFuture}
      onClick={handleClick}
      className="w-full card rounded-xl p-3.5 flex items-center justify-between gap-3 disabled:opacity-40 disabled:cursor-not-allowed transition-colors text-left"
    >
      <div className="flex items-center gap-2.5 min-w-0">
        <Avatar src={emp.avatar} name={emp.name} size={32} />
        <div className="min-w-0">
          <div className="text-[var(--text-primary)] text-sm font-medium truncate">{emp.name}</div>
          <div className="text-[var(--text-muted)] text-[11px]">{fmt(emp.dailyWage)}{t("perDay")}</div>
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

function ProfileDrawer({
  open, onClose, me, roleLabel, isAdmin, onDeleteAccount, onLogout,
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

      <div className="space-y-2">
        {groups.map((g) => {
          const isOpen = openKey === g.key;
          const groupTotal = g.items.reduce((sum, a) => sum + Number(a.amount), 0);
          return (
            <div key={g.key} className="rounded-lg overflow-hidden field">
              <button
                type="button"
                onClick={() => setOpenKey(isOpen ? null : g.key)}
                className="w-full flex items-center justify-between px-3.5 py-2.5"
              >
                <span className="text-[var(--text-primary)] text-xs font-semibold capitalize">{monthLabel(g.key)}</span>
                <div className="flex items-center gap-2">
                  <span className="text-[var(--text-muted)] text-[11px] font-mono tabular-nums">{fmt(groupTotal)}</span>
                  <ChevronDown size={14} className={`text-[var(--text-muted)] transition-transform ${isOpen ? "rotate-180" : ""}`} />
                </div>
              </button>
              {isOpen && (
                <div className="px-3.5 pb-1 bg-[var(--bg-card)]">
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

function AdminApp({
  usersData, currentUser, onLogout, summaryFor,
  adminTab, setAdminTab,
  newEmp, setNewEmp, empError, addEmployee, deleteEmployee, updateEmployeeWage, resetEmployeePassword,
  attendance, attDate, setAttDate, markAttendance, bulkMarkAttendance,
  advances, advEmp, setAdvEmp, advForm, setAdvForm, addAdvance, deleteAdvance,
  changeOwnCredentials, updateAvatar, deleteOwnAccount, accent, setAccent, mode, setMode, fontScale, setFontScale, lang, setLang, enableNotifications,
  notifications, markAllNotificationsRead, markNotificationRead, linkTelegram,
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
  // YANGI: "Ulashish" — brauzerning tabiiy ulashish oynasini ochadi (Web Share API).
  // Shu orqali foydalanuvchi hisobotni Telegram, WhatsApp, Gmail va hokazolarga
  // to'g'ridan-to'g'ri yubora oladi (fayl sifatida). SMS kabi fayl qabul qilmaydigan
  // kanallar uchun avtomatik matnli xulosaga tushadi. Agar qurilma/brauzer umuman
  // ulashishni qo'llab-quvvatlamasa (masalan kompyuter), oddiy Excel yuklab olishga tushadi.
  async function shareReport() {
    const wb = buildReportWorkbook();
    const fileName = `hisobot-${todayISO()}.xlsx`;
    try {
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const file = new File([blob], fileName, { type: blob.type });
      if (navigator.share && navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: t("reportHeader") });
        return;
      }
      if (navigator.share) {
        await navigator.share({ title: t("reportHeader"), text: buildReportTextSummary() });
        return;
      }
    } catch (e) {
      if (e && e.name === "AbortError") return; // foydalanuvchi ulashishni bekor qildi
    }
    XLSX.writeFile(wb, fileName);
  }
  const [weekOffset, setWeekOffset] = useState(0);
  const [attView, setAttView] = useState("daily"); // "daily" | "monthly"
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
            return (
              <>
          {visibleEmployees.length > 0 && attDate <= todayISO() && (
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
                    {pendingBulk.icon} {visibleEmployees.length} ta ishchiga "{pendingBulk.label}" qo'llansinmi?
                  </span>
                  <div className="flex gap-1.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => { bulkMarkAttendance(visibleEmployees.map((e) => e.id), pendingBulk.status); setPendingBulk(null); }}
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
                      <td className="py-2.5 px-5 text-[var(--text-primary)]">{emp.name}</td>
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
}) {
  const [drawerOpen, setDrawerOpen] = useState(false);
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
      <Shell
        title={t("employeePanel")}
        userName={currentUser.name}
        avatar={s.emp.avatar || null}
        onTitleClick={() => setDrawerOpen(true)}
        bottomNav={empBottomNav}
      >
        {empTab === "umumiy" && (
          <div className="tab-transition space-y-3">
            <div className="card rounded-xl p-6 text-center">
              <div className="flex items-center justify-center gap-1.5 text-[var(--text-muted)] text-xs mb-1.5">
                <Wallet size={13} /> {t("statRemainingSalary")}
              </div>
              <AnimatedAmount value={s.remaining} formatter={fmt} className="text-3xl font-bold font-mono tabular-nums text-[var(--good)]" />
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
  const [session, setSession] = useState(null);
  const [currentUser, setCurrentUserState] = useState(null);
  const [telegramPromptOpen, setTelegramPromptOpen] = useState(false);
  const [telegramPromptPhase, setTelegramPromptPhase] = useState("prompt"); // prompt | waiting | success
  const [telegramLinkBusy, setTelegramLinkBusy] = useState(false);
  const telegramPollRef = useRef(null);

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
      }

      setUsersData({
        admins: { [myProfile.username]: { avatar: myProfile.avatar_url } },
        employees,
      });
      setAttendance(attMap);
      setAdvances(advMap);
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
    if (!currentUser || currentUser.role !== "admin") return;
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
    if (!currentUser || currentUser.role !== "admin") return;
    await supabase.from("notifications").update({ is_read: true }).eq("admin_id", currentUser.id).eq("is_read", false);
    setNotifications((prev) => prev.map((n) => ({ ...n, is_read: true })));
  }

  async function markNotificationRead(id) {
    if (!currentUser || currentUser.role !== "admin") return;
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
      return;
    }
    setLoginBusy(true);
    try {
      const { data: email, error: lookupErr } = await supabase.rpc("email_for_username", { p_username: username });
      if (lookupErr || !email) {
        setLoginError(makeT(lang)("wrongLogin"));
        setLoginBusy(false);
        return;
      }
      const { data: signInData, error: signInErr } = await supabase.auth.signInWithPassword({ email, password });
      if (signInErr || !signInData?.session) {
        setLoginError(makeT(lang)("wrongLogin"));
        setLoginBusy(false);
        return;
      }
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

  function summaryFor(empId) {
    const emp = usersData.employees.find((x) => x.id === empId);
    if (!emp) return null;
    const att = attendance[empId] || {};
    let workedDays = 0;
    let totalWage = 0;
    for (const [date, raw] of Object.entries(att)) {
      const v = attEntryStatus(raw);
      workedDays += v;
      totalWage += v * attEntryWage(raw, emp, date);
    }
    const advList = advances[empId] || [];
    const totalAvans = advList.filter((a) => a.type !== "salary").reduce((sum, a) => sum + Number(a.amount), 0);
    const totalSalaryPaid = advList.filter((a) => a.type === "salary").reduce((sum, a) => sum + Number(a.amount), 0);
    const totalAdvance = totalAvans + totalSalaryPaid;
    return { emp, workedDays, totalWage, totalAdvance, totalAvans, totalSalaryPaid, remaining: totalWage - totalAdvance, advList, att };
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
      />
    );
  }

  return (
    <AppContext.Provider value={{ accent, lang, t }}>
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
