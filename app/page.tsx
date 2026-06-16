'use client';

import { supabase } from '@/lib/supabaseClient';
import type { Session } from '@supabase/supabase-js';
import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type ContractStatus =
  | 'Підготовка Комерційної пропозиції'
  | 'Тендерна процедура'
  | 'В процесі погодження'
  | 'Підготовка договору'
  | 'Договір на підписі'
  | 'Підписаний договір'
  | 'В процесі поставки товару'
  | 'Очікує оплати'
  | 'Сплачено'
  | 'Анульований договір';

type Contract = {
  id: string;
  contractNumber: string;
  title: string;
  customer: string;
  amount: number;
  status: ContractStatus;
  contractDate: string;
  deliveryDeadline: string;
  notes: string;
};
type ContractRow = {
  id: string;
  contract_number: string;
  title: string;
  customer: string;
  amount: number;
  status: ContractStatus;
  contract_date: string | null;
  delivery_deadline: string | null;
  notes: string | null;
  created_at?: string;
};
const STATUSES: ContractStatus[] = [
  'Підготовка Комерційної пропозиції',
  'Тендерна процедура',
  'В процесі погодження',
  'Підготовка договору',
  'Договір на підписі',
  'Підписаний договір',
  'В процесі поставки товару',
  'Очікує оплати',
  'Сплачено',
  'Анульований договір',
];

const emptyForm: Omit<Contract, 'id'> = {
  contractNumber: '',
  title: '',
  customer: '',
  amount: 0,
  status: 'Підготовка Комерційної пропозиції',
  contractDate: '',
  deliveryDeadline: '',
  notes: '',
};
function mapRowToContract(row: ContractRow): Contract {
  return {
    id: row.id,
    contractNumber: row.contract_number,
    title: row.title,
    customer: row.customer,
    amount: Number(row.amount),
    status: row.status,
    contractDate: row.contract_date || '',
    deliveryDeadline: row.delivery_deadline || '',
    notes: row.notes || '',
  };
}

function mapContractToRow(contract: Omit<Contract, 'id'>) {
  return {
    contract_number: contract.contractNumber,
    title: contract.title,
    customer: contract.customer,
    amount: contract.amount,
    status: contract.status,
    contract_date: contract.contractDate,
    delivery_deadline: contract.deliveryDeadline,
    notes: contract.notes,
  };
}
function formatMoney(value: number) {
  const formatted = new Intl.NumberFormat('uk-UA', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);

  return `${formatted} грн`;
}
function makeNextContractNumber(contracts: Contract[]) {
  const year = new Date().getFullYear();

  const numbers = contracts
    .map((item) => item.contractNumber)
    .map((number) => {
      const match = number.match(/^Д-(\d+)\/\d{4}$/);
      return match ? Number(match[1]) : 0;
    });

  const maxNumber = numbers.length ? Math.max(...numbers) : 0;
  const nextNumber = String(maxNumber + 1).padStart(3, '0');

  return `Д-${nextNumber}/${year}`;
}

export default function Home() {
  const [contracts, setContracts] = useState<Contract[]>([]);
  const [form, setForm] = useState<Omit<Contract, 'id'>>(emptyForm);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<'Всі статуси' | ContractStatus>('Всі статуси');
  const [isLoaded, setIsLoaded] = useState(false);

  const [session, setSession] = useState<Session | null>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');
const [showPassword, setShowPassword] = useState(false);
async function loadContracts() {
  const { data, error } = await supabase
    .from('contracts')
    .select('*')
    .order('created_at', { ascending: false });

  if (error) {
    console.error('Помилка завантаження договорів:', error);
    setIsLoaded(true);
    return;
  }

  setContracts((data || []).map((row) => mapRowToContract(row as ContractRow)));
  setIsLoaded(true);
}

useEffect(() => {
  async function initAuth() {
    const {
      data: { session },
    } = await supabase.auth.getSession();

    setSession(session);

    if (session) {
      await loadContracts();
    } else {
      setContracts([]);
      setIsLoaded(true);
    }

    setAuthLoading(false);
  }

  initAuth();

  const {
    data: { subscription },
  } = supabase.auth.onAuthStateChange((_event, newSession) => {
    setSession(newSession);

    if (newSession) {
      loadContracts();
    } else {
      setContracts([]);
    }
  });

  return () => {
    subscription.unsubscribe();
  };
}, []);
async function handleLogin() {
  setLoginError('');

  if (!loginEmail.trim() || !loginPassword.trim()) {
    setLoginError('Введіть email і пароль');
    return;
  }

  const { data, error } = await supabase.auth.signInWithPassword({
    email: loginEmail.trim(),
    password: loginPassword,
  });

  if (error) {
    console.error('Помилка входу:', error);
    setLoginError('Невірний email або пароль');
    return;
  }

  setSession(data.session);
  await loadContracts();
}

async function handleLogout() {
  await supabase.auth.signOut();
  setSession(null);
  setContracts([]);
}
  const filteredContracts = useMemo(() => {
    return contracts.filter((contract) => {
      const text = `${contract.contractNumber} ${contract.title} ${contract.customer} ${contract.notes}`.toLowerCase();

      const matchesSearch = text.includes(search.toLowerCase());
      const matchesStatus =
        statusFilter === 'Всі статуси' || contract.status === statusFilter;

      return matchesSearch && matchesStatus;
    });
  }, [contracts, search, statusFilter]);

  const totalAmount = useMemo(() => {
    return filteredContracts.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  }, [filteredContracts]);

  const paidAmount = useMemo(() => {
    return contracts
      .filter((item) => item.status === 'Сплачено')
      .reduce((sum, item) => sum + Number(item.amount || 0), 0);
  }, [contracts]);

  const unpaidAmount = useMemo(() => {
    return contracts
      .filter((item) => item.status !== 'Сплачено' && item.status !== 'Анульований договір')
      .reduce((sum, item) => sum + Number(item.amount || 0), 0);
  }, [contracts]);

  function handleChange(
    field: keyof Omit<Contract, 'id'>,
    value: string | number
  ) {
    setForm((prev) => ({
      ...prev,
      [field]: value,
    }));
  }

  async function handleSubmit() {
  if (!form.title.trim()) {
    alert('Вкажіть назву договору або предмет закупівлі');
    return;
  }

  if (!form.amount || Number(form.amount) <= 0) {
    alert('Вкажіть суму договору');
    return;
  }

  const preparedForm: Omit<Contract, 'id'> = {
    ...form,
    contractNumber: form.contractNumber || makeNextContractNumber(contracts),
    amount: Number(form.amount),
  };

  if (editingId) {
    const { data, error } = await supabase
      .from('contracts')
      .update(mapContractToRow(preparedForm))
      .eq('id', editingId)
      .select()
      .single();

    if (error) {
      console.error('Помилка оновлення договору:', error);
      alert('Не вдалося оновити договір');
      return;
    }

    setContracts((prev) =>
      prev.map((item) =>
        item.id === editingId ? mapRowToContract(data as ContractRow) : item
      )
    );

    setEditingId(null);
    setForm(emptyForm);
    return;
  }

  const { data, error } = await supabase
    .from('contracts')
    .insert(mapContractToRow(preparedForm))
    .select()
    .single();

  if (error) {
    console.error('Помилка створення договору:', error);
    alert('Не вдалося створити договір');
    return;
  }

  setContracts((prev) => [mapRowToContract(data as ContractRow), ...prev]);
  setForm(emptyForm);
}

  function handleEdit(contract: Contract) {
    setEditingId(contract.id);

    setForm({
      contractNumber: contract.contractNumber,
      title: contract.title,
      customer: contract.customer,
      amount: contract.amount,
      status: contract.status,
      contractDate: contract.contractDate,
      deliveryDeadline: contract.deliveryDeadline,
      notes: contract.notes,
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  async function handleDelete(id: string) {
  const confirmed = confirm('Точно видалити цей договір?');

  if (!confirmed) return;

  const { error } = await supabase
    .from('contracts')
    .delete()
    .eq('id', id);

  if (error) {
    console.error('Помилка видалення договору:', error);
    alert('Не вдалося видалити договір');
    return;
  }

  setContracts((prev) => prev.filter((item) => item.id !== id));
}
async function handleStatusChange(id: string, status: ContractStatus) {
  const { data, error } = await supabase
    .from('contracts')
    .update({ status })
    .eq('id', id)
    .select()
    .single();

  if (error) {
    console.error('Помилка зміни статусу:', error);
    alert('Не вдалося змінити статус');
    return;
  }

  setContracts((prev) =>
    prev.map((item) =>
      item.id === id ? mapRowToContract(data as ContractRow) : item
    )
  );
}  function handleCancelEdit() {
    setEditingId(null);
    setForm(emptyForm);
  }

  function exportToCsv() {
    const headers = [
      '№ договору',
      'Предмет закупівлі',
      'Замовник',
      'Сума',
      'Статус',
      'Дата договору',
      'Дедлайн поставки',
      'Примітки',
    ];

    const rows = contracts.map((item) => [
      item.contractNumber,
      item.title,
      item.customer,
      item.amount,
      item.status,
      item.contractDate,
      item.deliveryDeadline,
      item.notes,
    ]);

    const csv = [headers, ...rows]
      .map((row) =>
        row
          .map((cell) => `"${String(cell).replaceAll('"', '""')}"`)
          .join(';')
      )
      .join('\n');

    const blob = new Blob(['\uFEFF' + csv], {
      type: 'text/csv;charset=utf-8;',
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = 'contracts-crm.csv';
    link.click();

    URL.revokeObjectURL(url);
  }
function exportToExcel() {
  const rows = contracts.map((item) => ({
    '№ договору': item.contractNumber,
    'Предмет закупівлі': item.title,
    'Замовник': item.customer,
    'Сума': item.amount,
    'Статус': item.status,
    'Дата договору': item.contractDate,
    'Дедлайн поставки': item.deliveryDeadline,
    'Примітки': item.notes,
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);

  worksheet['!cols'] = [
    { wch: 16 },
    { wch: 45 },
    { wch: 28 },
    { wch: 16 },
    { wch: 32 },
    { wch: 16 },
    { wch: 18 },
    { wch: 45 },
  ];

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, 'Договори');

  XLSX.writeFile(workbook, 'contracts-crm.xlsx');
}
if (authLoading) {
  return (
    <main className="page authPage">
      <section className="authCard">
        <h1>CRM облік договорів</h1>
        <p>Завантаження...</p>
      </section>
    </main>
  );
}

if (!session) {
  return (
    <main className="page authPage">
      <section className="authCard">
        <h1>CRM облік договорів</h1>
        <p>Увійдіть, щоб працювати з договорами.</p>

        <div className="authForm">
          <label>
            Email
            <input
              type="email"
              value={loginEmail}
              autoComplete="username"
              placeholder="email@example.com"
              onChange={(e) => setLoginEmail(e.target.value)}
            />
          </label>

        <label>
  Пароль

  <div className="passwordField">
    <input
      type={showPassword ? 'text' : 'password'}
      value={loginPassword}
      autoComplete="current-password"
      placeholder="Ваш пароль"
      onChange={(e) => setLoginPassword(e.target.value)}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          handleLogin();
        }
      }}
    />

    <button
      type="button"
      className="passwordToggle"
      onClick={() => setShowPassword((prev) => !prev)}
    >
      {showPassword ? 'Сховати' : 'Показати'}
    </button>
  </div>
</label>

          {loginError && <div className="authError">{loginError}</div>}

          <button className="primaryButton" onClick={handleLogin}>
            Увійти
          </button>

          <p className="authHint">
            На iPhone збережіть пароль в iCloud Keychain — наступного разу Safari
            або Chrome запропонує вхід через Face ID.
          </p>
        </div>
      </section>
    </main>
  );
}
  return (
    <main className="page">
      <section className="header">
        <div>
          <h1>CRM облік договорів</h1>
          <p>Система для контролю договорів, поставок, оплат і статусів.</p>
        </div>

   <div className="exportButtons">
  <button className="secondaryButton" onClick={exportToCsv}>
    Експорт CSV
  </button>

  <button className="secondaryButton" onClick={exportToExcel}>
    Експорт Excel
  </button>

  <button className="secondaryButton" onClick={handleLogout}>
    Вийти
  </button>
</div>
      </section>

      <section className="statsGrid">
        <div className="statCard">
          <span>Кількість договорів</span>
          <strong>{contracts.length}</strong>
        </div>

        <div className="statCard">
          <span>Сума у фільтрі</span>
          <strong>{formatMoney(totalAmount)}</strong>
        </div>

        <div className="statCard">
          <span>Сплачено</span>
          <strong>{formatMoney(paidAmount)}</strong>
        </div>

        <div className="statCard">
          <span>Очікується</span>
          <strong>{formatMoney(unpaidAmount)}</strong>
        </div>
      </section>

      <section className="card">
        <h2>{editingId ? 'Редагування договору' : 'Новий договір'}</h2>

        <div className="formGrid">
          <label>
            № договору
            <input
              value={form.contractNumber}
              placeholder="Автоматично, якщо залишити пустим"
              onChange={(e) => handleChange('contractNumber', e.target.value)}
            />
          </label>

          <label>
            Предмет закупівлі / назва договору
            <input
              value={form.title}
              onChange={(e) => handleChange('title', e.target.value)}
            />
          </label>

          <label>
            Замовник / контрагент
            <input
              value={form.customer}
              onChange={(e) => handleChange('customer', e.target.value)}
            />
          </label>

          <label>
            Сума договору, грн
            <input
              type="number"
              value={form.amount || ''}
              placeholder="0.00"
              onChange={(e) => handleChange('amount', Number(e.target.value))}
            />
          </label>

          <label>
            Статус
            <select
              value={form.status}
              onChange={(e) =>
                handleChange('status', e.target.value as ContractStatus)
              }
            >
              {STATUSES.map((status) => (
                <option key={status} value={status}>
                  {status}
                </option>
              ))}
            </select>
          </label>

          <label>
            Дата договору
            <input
              type="date"
              value={form.contractDate}
              onChange={(e) => handleChange('contractDate', e.target.value)}
            />
          </label>

          <label>
            Дедлайн поставки
            <input
              type="date"
              value={form.deliveryDeadline}
              onChange={(e) => handleChange('deliveryDeadline', e.target.value)}
            />
          </label>

          <label className="wide">
            Примітки
            <textarea
              value={form.notes}
              onChange={(e) => handleChange('notes', e.target.value)}
            />
          </label>
        </div>

        <div className="actions">
          <button className="primaryButton" onClick={handleSubmit}>
            {editingId ? 'Зберегти зміни' : 'Додати договір'}
          </button>

          {editingId && (
            <button className="secondaryButton" onClick={handleCancelEdit}>
              Скасувати редагування
            </button>
          )}
        </div>
      </section>

      <section className="card">
        <div className="tableHeader">
          <h2>Список договорів</h2>

          <div className="filters">
            <input
              value={search}
              placeholder="Пошук..."
              onChange={(e) => setSearch(e.target.value)}
            />

            <select
              value={statusFilter}
              onChange={(e) =>
                setStatusFilter(e.target.value as 'Всі статуси' | ContractStatus)
              }
            >
              <option value="Всі статуси">Всі статуси</option>

              {STATUSES.map((status) => (
                <option key={status} value={status}>
                  {status}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="desktopOnly tableWrapper">
  <table>
    <thead>
      <tr>
        <th>№</th>
        <th>Предмет закупівлі</th>
        <th>Замовник</th>
        <th>Сума</th>
        <th>Статус</th>
        <th>Дата</th>
        <th>Дедлайн</th>
        <th>Примітки</th>
        <th>Дії</th>
      </tr>
    </thead>

    <tbody>
      {filteredContracts.length === 0 ? (
        <tr>
          <td colSpan={9} className="empty">
            Поки немає договорів
          </td>
        </tr>
      ) : (
        filteredContracts.map((contract) => (
          <tr key={contract.id}>
            <td>{contract.contractNumber}</td>
            <td className="titleCell">{contract.title}</td>
            <td>{contract.customer || '—'}</td>
            <td>{formatMoney(contract.amount)}</td>
            <td>
              <select
                className={`statusSelect ${getStatusClass(contract.status)}`}
                value={contract.status}
                onChange={(e) =>
                  handleStatusChange(contract.id, e.target.value as ContractStatus)
                }
              >
                {STATUSES.map((status) => (
                  <option key={status} value={status}>
                    {status}
                  </option>
                ))}
              </select>
            </td>
            <td>{contract.contractDate || '—'}</td>
            <td>{contract.deliveryDeadline || '—'}</td>
            <td>{contract.notes || '—'}</td>
            <td>
              <div className="rowActions">
                <button onClick={() => handleEdit(contract)}>
                  Редагувати
                </button>
                <button onClick={() => handleDelete(contract.id)}>
                  Видалити
                </button>
              </div>
            </td>
          </tr>
        ))
      )}
    </tbody>
  </table>
</div>

<div className="mobileOnly mobileContractsList">
  {filteredContracts.length === 0 ? (
    <div className="mobileEmpty">Поки немає договорів</div>
  ) : (
    filteredContracts.map((contract) => (
      <article className="contractMobileCard" key={contract.id}>
        <div className="contractMobileTop">
          <span className="contractMobileNumber">
            {contract.contractNumber}
          </span>

          <strong className="contractMobileAmount">
            {formatMoney(contract.amount)}
          </strong>
        </div>

        <h3 className="contractMobileTitle">{contract.title}</h3>

        <div className="contractMobileInfo">
          <div>
            <span>Замовник</span>
            <strong>{contract.customer || '—'}</strong>
          </div>

          <div>
            <span>Дата договору</span>
            <strong>{contract.contractDate || '—'}</strong>
          </div>

          <div>
            <span>Дедлайн</span>
            <strong>{contract.deliveryDeadline || '—'}</strong>
          </div>
        </div>

        <div className="contractMobileStatus">
          <span>Статус</span>

          <select
            className={`statusSelect ${getStatusClass(contract.status)}`}
            value={contract.status}
            onChange={(e) =>
              handleStatusChange(contract.id, e.target.value as ContractStatus)
            }
          >
            {STATUSES.map((status) => (
              <option key={status} value={status}>
                {status}
              </option>
            ))}
          </select>
        </div>

        {contract.notes && (
          <div className="contractMobileNotes">
            <span>Примітки</span>
            <p>{contract.notes}</p>
          </div>
        )}

        <div className="contractMobileActions">
          <button onClick={() => handleEdit(contract)}>
            Редагувати
          </button>

          <button onClick={() => handleDelete(contract.id)}>
            Видалити
          </button>
        </div>
      </article>
    ))
  )}
</div>
      </section>
    </main>
  );
}

function getStatusClass(status: ContractStatus) {
  switch (status) {
    case 'Сплачено':
      return 'paid';
    case 'Анульований договір':
      return 'cancelled';
    case 'Очікує оплати':
      return 'waiting';
    case 'В процесі поставки товару':
      return 'delivery';
    case 'Підписаний договір':
      return 'signed';
    default:
      return 'progress';
  }
}