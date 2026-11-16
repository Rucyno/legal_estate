import React, { useState } from 'react';
import { motion, AnimatePresence, useReducedMotion } from 'framer-motion';
import {
  PropertyType,
  PartyComposition,
  EncumbranceStatus,
  PaymentMethod,
  FinancingSource,
  DealScenarioResult
} from './DealMatrixTypes';

export const FigJamDealMatrix: React.FC = () => {
  const shouldReduceMotion = useReducedMotion();
  const [activeTab, setActiveTab] = useState<'simulator' | 'matrix' | 'contract_blocks' | 'prep_checklist'>('simulator');

  // Simulator State
  const [property, setProperty] = useState<PropertyType>('apartment');
  const [seller, setSeller] = useState<PartyComposition>('adult_1');
  const [buyer, setBuyer] = useState<PartyComposition>('spouses');
  const [encumbrance, setEncumbrance] = useState<EncumbranceStatus>('free');
  const [financing, setFinancing] = useState<FinancingSource>('own_plus_mortgage');
  const [paymentMethod, setPaymentMethod] = useState<PaymentMethod>('sbr');

  // Matrix Filter State
  const [matrixBranch, setMatrixBranch] = useState<'free' | 'mortgaged'>('free');
  const [selectedNodeId, setSelectedNodeId] = useState<string | null>(null);

  // Animation variants
  const fadeIn = {
    initial: { opacity: 0, y: shouldReduceMotion ? 0 : 12 },
    animate: { opacity: 1, y: 0, transition: { duration: 0.28, ease: 'easeOut' } },
    exit: { opacity: 0, y: shouldReduceMotion ? 0 : -8, transition: { duration: 0.2 } }
  };

  const cardVariant = {
    initial: { opacity: 0, scale: shouldReduceMotion ? 1 : 0.98 },
    animate: { opacity: 1, scale: 1, transition: { duration: 0.25 } },
    hover: { y: shouldReduceMotion ? 0 : -3, transition: { duration: 0.15 } }
  };

  // Calculation logic based on FigJam settlement matrix
  const calculateDealScenario = (): DealScenarioResult => {
    const isFree = encumbrance === 'free';
    const hasMortgage = financing === 'own_plus_mortgage' || financing === 'combined_all';
    const hasMatcap = financing === 'own_plus_matcap' || financing === 'combined_all';

    let title = '';
    let encumbranceRule = '';
    let statuteArticle = '';
    let matcapRule: string | undefined = undefined;
    let receiptRule = '';
    const steps: string[] = [];
    const risksAndWarnings: string[] = [];
    let recommendedContractClause = '';

    // Step 1: Base Title
    if (isFree) {
      if (financing === 'own_funds') {
        title = 'Собственные средства 100% (Объект без залога)';
        statuteArticle = paymentMethod === 'cash_full'
          ? 'п. 5 ст. 488 ГК РФ — Залог продавца не возникает (деньги переданы до ОДКП)'
          : 'ст. 488 ГК РФ — Без залога продавца (расчет после регистрации через безопасный инструмент)';
      } else if (financing === 'own_plus_matcap') {
        title = 'Собственные средства + Материнский капитал (Без ипотеки)';
        statuteArticle = 'ст. 10 Федерального закона № 256-ФЗ, ст. 488 ГК РФ (залог продавца до поступления средств СФР исключается по соглашению сторон)';
      } else if (financing === 'own_plus_mortgage') {
        title = 'Собственные средства + Ипотека банка (Объект без залога)';
        statuteArticle = 'ст. 77 Федерального закона № 102-ФЗ «Об ипотеке» (Ипотека в силу закона в пользу Банка). Ст. 488 ГК РФ строго НЕ применяется!';
      } else {
        title = 'Комбинированная сделка: Собственные + Ипотека + МСК';
        statuteArticle = 'ст. 77 102-ФЗ (Ипотека в пользу Банка), ст. 10 256-ФЗ (МСК). Залог продавца по ст. 488 ГК РФ исключен!';
      }
    } else {
      // Mortgaged Seller
      if (financing === 'own_funds') {
        title = 'Покупка за собственные средства (Объект в залоге у банка Продавца)';
        statuteArticle = 'ст. 353 ГК РФ, регламент залогодержателя. Залог продавца по ст. 488 ГК РФ не возникает.';
      } else if (financing === 'own_plus_mortgage') {
        title = 'Сделка «Ипотека на Ипотеку» (Обременение Продавца + Кредит Покупателя)';
        statuteArticle = 'ст. 77 102-ФЗ, ст. 313 ГК РФ, регламент межбанковской/внутрибанковской переипотеки. Залог продавца исключен.';
      } else if (financing === 'own_plus_matcap') {
        title = 'Залоговый объект + Собственные средства + МСК (Без ипотеки покупателя)';
        statuteArticle = 'ст. 10 256-ФЗ. Залог банка гасится ТОЛЬКО личными средствами покупателя до или на сделке!';
      } else {
        title = 'Залоговый объект + Ипотека Покупателя + МСК (Многосоставная сделка)';
        statuteArticle = 'ст. 77 102-ФЗ, ст. 10 256-ФЗ, ст. 313 ГК РФ.';
      }
    }

    // Step 2: Encumbrance Rule & Statute
    if (hasMortgage) {
      encumbranceRule = 'Ипотека в силу закона в пользу Банка-кредитора. Право залога Продавца по п. 5 ст. 488 ГК РФ не возникает!';
    } else {
      encumbranceRule = 'Объект не обременяется залогом кредитной организации.';
    }

    // Step 3: Matcap Rule
    if (hasMatcap) {
      if (hasMortgage) {
        matcapRule = 'Возраст ребенка НЕ ограничивается 3 годами (ч. 6 ст. 10 256-ФЗ — допускается использование МСК в любой момент при кредитовании).';
      } else {
        matcapRule = 'КРИТИЧНО: Ребенку, давшему право на сертификат МСК, на момент заключения ДКП должно быть СТРОГО БОЛЕЕ 3 ЛЕТ!';
      }
    }

    // Step 4: Receipt Rule
    if (paymentMethod === 'cash_full') {
      receiptRule = 'Общая расписка на полную сумму стоимости объекта с учетом задатка по ПДКП (составляется и подписывается в день заключения ОДКП).';
    } else if (paymentMethod === 'sbr') {
      receiptRule = 'В день сделки выдается расписка ТОЛЬКО на переданный задаток. На сумму в СБР расписка выдается после фактического зачисления средств на счет продавца.';
    } else if (paymentMethod === 'letter_of_credit') {
      receiptRule = 'В день сделки — расписка строго на сумму задатка. Средства на аккредитиве еще не получены, расписка на остаток подписывается при раскрытии.';
    } else {
      receiptRule = 'В день сделки — расписка на задаток. На сумму в ячейке составляется гарантийное обязательство/заявление; расписка оформляется при выемке из ячейки.';
    }

    // Step 5: Sequential Steps
    steps.push('1. Заключение Предварительного договора (ПДКП) с фиксацией задатка/обеспечительного платежа.');
    if (!isFree) {
      steps.push('2. Заказ справки об остатке ссудной задолженности продавца и согласование технологии снятия обременения с банком-залогодержателем.');
    }
    if (paymentMethod === 'sbr') {
      steps.push('3. Открытие номинального счета СБР и внесение денежных средств покупателем.');
    } else if (paymentMethod === 'letter_of_credit') {
      steps.push('3. Открытие безотзывного покрытого' + (!isFree ? ' трансферабельного' : '') + ' аккредитива с депонированием средств.');
    } else if (paymentMethod === 'safe_deposit_box') {
      steps.push('3. Закладка денежных средств в банковский индивидуальный сейф (ячейку) с условием доступа по выписке ЕГРН.');
    } else {
      steps.push('3. Передача остатка наличных денежных средств до подписания ОДКП.');
    }
    steps.push('4. Подписание Основного договора (ОДКП) и подача пакета на государственную регистрацию в Росреестр.');
    steps.push('5. Переход права собственности в ЕГРН и раскрытие денежных средств Продавцу (1-5 рабочих дней).');
    if (hasMatcap) {
      steps.push('6. Подача заявления в Социальный фонд России (СФР) на распоряжение МСК; выплата продавцу в срок до 21 дня.');
    }

    // Step 6: Risks and Specific Warnings
    if (property === 'room') {
      risksAndWarnings.push('Ст. 250 ГК РФ: Обязательны нотариальные отказы от преимущественного права покупки от всех соседей либо отправка уведомлений через нотариуса за 30 дней.');
    }
    if (property === 'apartment_share') {
      risksAndWarnings.push('Сделки по отчуждению долей в праве общей собственности подлежат обязательному нотариальному удостоверению (ст. 42 218-ФЗ).');
    }
    if (seller === 'minor_under_14' || seller === 'minor_14_to_18') {
      risksAndWarnings.push('Обязательно предварительное разрешение Органов опеки и попечительства (ст. 37 ГК РФ) с одновременным выделением встречной доли/жилья.');
    }
    if (seller === 'spouses') {
      risksAndWarnings.push('Обязательно нотариальное согласие супруга(и) Продавца (ст. 35 СК РФ) либо брачный договор.');
    }
    if (!isFree && paymentMethod === 'cash_full') {
      risksAndWarnings.push('КРИТИЧЕСКИЙ РИСК: Передача наличных на гашение ипотеки Продавца до сделки без банковского контроля. Требует целевого займа и безотзывной доверенности.');
    }

    recommendedContractClause = hasMortgage
      ? '«В соответствии со ст. 77 Федерального закона "Об ипотеке (залоге недвижимости)" с момента государственной регистрации права собственности Покупателя Объект считается находящимся в залоге у Банка. В силу п. 5 ст. 488 ГК РФ право залога у Продавца не возникает.»'
      : '«В соответствии с п. 5 ст. 488 Гражданского кодекса РФ право залога у Продавца на Объект недвижимости до момента полной оплаты не возникает по взаимному соглашению Сторон.»';

    return {
      title,
      encumbranceRule,
      statuteArticle,
      matcapRule,
      receiptRule,
      steps,
      risksAndWarnings,
      recommendedContractClause
    };
  };

  const dealResult = calculateDealScenario();

  return (
    <div style={{
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
      backgroundColor: '#F8FAFC',
      color: '#0F172A',
      minHeight: '100vh',
      padding: '24px',
      boxSizing: 'border-box'
    }}>
      <div style={{ maxWidth: 1320, margin: '0 auto' }}>
        
        {/* Brand Header */}
        <motion.header
          initial={{ opacity: 0, y: -16 }}
          animate={{ opacity: 1, y: 0 }}
          style={{
            backgroundColor: '#FFFFFF',
            borderRadius: '14px',
            padding: '20px 28px',
            border: '1px solid #E2E8F0',
            borderTop: '5px solid #DC2626',
            boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)',
            marginBottom: '24px',
            display: 'flex',
            flexWrap: 'wrap',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '16px'
          }}
        >
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '4px' }}>
              <span style={{
                backgroundColor: '#DC2626',
                color: '#FFFFFF',
                fontWeight: 800,
                fontSize: '12px',
                padding: '4px 10px',
                borderRadius: '6px',
                letterSpacing: '0.5px'
              }}>
                ЭТАЖИ ЮРИСТ
              </span>
              <span style={{ fontSize: '13px', color: '#64748B', fontWeight: 500 }}>
                FigJam Settlement & Legal Engine v2.0
              </span>
            </div>
            <h1 style={{ fontSize: '24px', fontWeight: 700, color: '#0F172A', margin: 0 }}>
              Интерактивная карта сделок и матрица расчетов
            </h1>
            <p style={{ fontSize: '13px', color: '#64748B', margin: '4px 0 0 0' }}>
              Интеграция корпоративной доски FigJam «Welcome to FigJam»: правовые режимы, ст. 488 ГК РФ, 77 102-ФЗ, регламенты расписок и МСК.
            </p>
          </div>

          {/* Navigation Tabs */}
          <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
            {[
              { id: 'simulator', label: '⚡ Симулятор сделки' },
              { id: 'matrix', label: '🗺️ Карта расчетов FigJam' },
              { id: 'contract_blocks', label: '📄 Блочная структура ПДКП' },
              { id: 'prep_checklist', label: '📋 Чек-лист подготовки' }
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id as any)}
                style={{
                  padding: '9px 16px',
                  borderRadius: '8px',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  border: 'none',
                  backgroundColor: activeTab === tab.id ? '#DC2626' : '#F1F5F9',
                  color: activeTab === tab.id ? '#FFFFFF' : '#475569',
                  transition: 'all 0.2s ease'
                }}
              >
                {tab.label}
              </button>
            ))}
          </div>
        </motion.header>

        {/* TAB 1: DEAL SIMULATOR */}
        <AnimatePresence mode="wait">
          {activeTab === 'simulator' && (
            <motion.div key="simulator" {...fadeIn}>
              <div style={{ display: 'grid', gridTemplateColumns: 'minmax(320px, 440px) 1fr', gap: '24px' }}>
                
                {/* Control Panel */}
                <div style={{
                  backgroundColor: '#FFFFFF',
                  borderRadius: '12px',
                  padding: '24px',
                  border: '1px solid #E2E8F0',
                  boxShadow: '0 2px 4px rgba(0,0,0,0.03)'
                }}>
                  <h2 style={{ fontSize: '16px', fontWeight: 700, marginBottom: '18px', color: '#0F172A' }}>
                    Параметры сделки
                  </h2>

                  {/* Property Selector */}
                  <div style={{ marginBottom: '16px' }}>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Объект недвижимости
                    </label>
                    <select
                      value={property}
                      onChange={(e) => setProperty(e.target.value as PropertyType)}
                      style={{
                        width: '100%',
                        padding: '9px 12px',
                        borderRadius: '6px',
                        border: '1px solid #CBD5E1',
                        fontSize: '13px',
                        color: '#0F172A',
                        backgroundColor: '#FFFFFF',
                        outline: 'none'
                      }}
                    >
                      <option value="apartment">Квартира (целиком)</option>
                      <option value="apartment_share">Доля в квартире (нотариат)</option>
                      <option value="house_and_land">Жилой дом и земельный участок</option>
                      <option value="land">Земельный участок</option>
                      <option value="room">Комната (ст. 250 ГК РФ)</option>
                      <option value="commercial">Нежилое коммерческое помещение</option>
                      <option value="land_and_outbuilding">Земельный участок + гараж/баня</option>
                    </select>
                  </div>

                  {/* Seller Composition */}
                  <div style={{ marginBottom: '16px' }}>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Сторона Продавца
                    </label>
                    <select
                      value={seller}
                      onChange={(e) => setSeller(e.target.value as PartyComposition)}
                      style={{
                        width: '100%',
                        padding: '9px 12px',
                        borderRadius: '6px',
                        border: '1px solid #CBD5E1',
                        fontSize: '13px',
                        color: '#0F172A',
                        backgroundColor: '#FFFFFF'
                      }}
                    >
                      <option value="adult_1">1 взрослый собственник</option>
                      <option value="adult_2">2 взрослых собственника</option>
                      <option value="adult_3">3 взрослых собственника</option>
                      <option value="spouses">2 взрослых (супруги, ст. 35 СК РФ)</option>
                      <option value="minor_under_14">Собственник до 14 лет (Опека)</option>
                      <option value="minor_14_to_18">Собственник от 14 до 18 лет (Опека)</option>
                    </select>
                  </div>

                  {/* Buyer Composition */}
                  <div style={{ marginBottom: '16px' }}>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Сторона Покупателя
                    </label>
                    <select
                      value={buyer}
                      onChange={(e) => setBuyer(e.target.value as PartyComposition)}
                      style={{
                        width: '100%',
                        padding: '9px 12px',
                        borderRadius: '6px',
                        border: '1px solid #CBD5E1',
                        fontSize: '13px',
                        color: '#0F172A',
                        backgroundColor: '#FFFFFF'
                      }}
                    >
                      <option value="adult_1">1 взрослый покупатель</option>
                      <option value="spouses">Супруги (совместная собственность)</option>
                      <option value="adult_2">2 взрослых покупателя (доли)</option>
                      <option value="minor_under_14">Покупатель до 14 лет (ребенок)</option>
                    </select>
                  </div>

                  {/* Encumbrance status */}
                  <div style={{ marginBottom: '16px' }}>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Обременение объекта (Ипотека Продавца)
                    </label>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                      <button
                        onClick={() => setEncumbrance('free')}
                        style={{
                          padding: '10px',
                          borderRadius: '6px',
                          border: encumbrance === 'free' ? '2px solid #16A34A' : '1px solid #CBD5E1',
                          backgroundColor: encumbrance === 'free' ? '#F0FDF4' : '#FFFFFF',
                          color: encumbrance === 'free' ? '#16A34A' : '#475569',
                          fontWeight: 600,
                          fontSize: '12px',
                          cursor: 'pointer'
                        }}
                      >
                        ✅ Чистый (Без залога)
                      </button>
                      <button
                        onClick={() => setEncumbrance('mortgaged')}
                        style={{
                          padding: '10px',
                          borderRadius: '6px',
                          border: encumbrance === 'mortgaged' ? '2px solid #DC2626' : '1px solid #CBD5E1',
                          backgroundColor: encumbrance === 'mortgaged' ? '#FEF2F2' : '#FFFFFF',
                          color: encumbrance === 'mortgaged' ? '#DC2626' : '#475569',
                          fontWeight: 600,
                          fontSize: '12px',
                          cursor: 'pointer'
                        }}
                      >
                        ⚠️ В залоге банка
                      </button>
                    </div>
                  </div>

                  {/* Financing Source */}
                  <div style={{ marginBottom: '16px' }}>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Источник средств Покупателя
                    </label>
                    <select
                      value={financing}
                      onChange={(e) => setFinancing(e.target.value as FinancingSource)}
                      style={{
                        width: '100%',
                        padding: '9px 12px',
                        borderRadius: '6px',
                        border: '1px solid #CBD5E1',
                        fontSize: '13px',
                        color: '#0F172A',
                        backgroundColor: '#FFFFFF'
                      }}
                    >
                      <option value="own_funds">Собственные средства 100%</option>
                      <option value="own_plus_matcap">Собственные + МСК (без ипотеки)</option>
                      <option value="own_plus_mortgage">Собственные + Ипотека банка</option>
                      <option value="combined_all">Собственные + Ипотека + МСК</option>
                    </select>
                  </div>

                  {/* Payment Method */}
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: '#475569', marginBottom: '6px' }}>
                      Инструмент передачи денег
                    </label>
                    <select
                      value={paymentMethod}
                      onChange={(e) => setPaymentMethod(e.target.value as PaymentMethod)}
                      style={{
                        width: '100%',
                        padding: '9px 12px',
                        borderRadius: '6px',
                        border: '1px solid #CBD5E1',
                        fontSize: '13px',
                        color: '#0F172A',
                        backgroundColor: '#FFFFFF'
                      }}
                    >
                      <option value="sbr">СБР (Система безопасных расчетов)</option>
                      <option value="letter_of_credit">Аккредитив (покрытый безотзывный)</option>
                      <option value="cash_full">Наличные расчеты</option>
                      <option value="safe_deposit_box">Банковская сейфовая ячейка</option>
                    </select>
                  </div>
                </div>

                {/* Live Legal Output Card */}
                <motion.div
                  layout
                  style={{
                    backgroundColor: '#FFFFFF',
                    borderRadius: '12px',
                    padding: '28px',
                    border: '1px solid #E2E8F0',
                    boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '20px'
                  }}
                >
                  <div>
                    <span style={{
                      backgroundColor: encumbrance === 'free' ? '#DC2626' : '#B91C1C',
                      color: '#FFFFFF',
                      fontSize: '11px',
                      fontWeight: 700,
                      padding: '3px 8px',
                      borderRadius: '4px',
                      textTransform: 'uppercase'
                    }}>
                      {encumbrance === 'free' ? 'СТВОЛ I: Объект чистый' : 'СТВОЛ II: Залог кредитора'}
                    </span>
                    <h2 style={{ fontSize: '20px', fontWeight: 700, color: '#0F172A', margin: '8px 0 4px 0' }}>
                      {dealResult.title}
                    </h2>
                    <p style={{ fontSize: '13px', color: '#475569', margin: 0, fontWeight: 500 }}>
                      Основание: {dealResult.statuteArticle}
                    </p>
                  </div>

                  {/* Critical Alert Boxes */}
                  {dealResult.matcapRule && (
                    <div style={{
                      backgroundColor: financing === 'own_plus_matcap' ? '#FEF2F2' : '#F0FDF4',
                      borderLeft: `4px solid ${financing === 'own_plus_matcap' ? '#DC2626' : '#16A34A'}`,
                      padding: '12px 16px',
                      borderRadius: '6px',
                      fontSize: '13px',
                      fontWeight: 600,
                      color: financing === 'own_plus_matcap' ? '#991B1B' : '#166534'
                    }}>
                      🍼 ПРАВИЛО МСК: {dealResult.matcapRule}
                    </div>
                  )}

                  {dealResult.risksAndWarnings.length > 0 && (
                    <div style={{
                      backgroundColor: '#FFFBEB',
                      borderLeft: '4px solid #D97706',
                      padding: '12px 16px',
                      borderRadius: '6px'
                    }}>
                      <div style={{ fontSize: '12px', fontWeight: 700, color: '#B45309', marginBottom: '4px' }}>
                        ⚠️ ЮРИДИЧЕСКИЕ ПРЕДУПРЕЖДЕНИЯ И РИСКИ:
                      </div>
                      <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: '#92400E' }}>
                        {dealResult.risksAndWarnings.map((warn, idx) => (
                          <li key={idx} style={{ marginBottom: '2px' }}>{warn}</li>
                        ))}
                      </ul>
                    </div>
                  )}

                  {/* Statute & Receipts */}
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
                    <div style={{
                      backgroundColor: '#F8FAFC',
                      padding: '14px',
                      borderRadius: '8px',
                      border: '1px solid #E2E8F0'
                    }}>
                      <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748B', textTransform: 'uppercase', marginBottom: '4px' }}>
                        Статус залога продавца
                      </div>
                      <div style={{ fontSize: '13px', color: '#0F172A', fontWeight: 600 }}>
                        {dealResult.encumbranceRule}
                      </div>
                    </div>

                    <div style={{
                      backgroundColor: '#F8FAFC',
                      padding: '14px',
                      borderRadius: '8px',
                      border: '1px solid #E2E8F0'
                    }}>
                      <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748B', textTransform: 'uppercase', marginBottom: '4px' }}>
                        Регламент расписок
                      </div>
                      <div style={{ fontSize: '13px', color: '#0F172A', fontWeight: 600 }}>
                        {dealResult.receiptRule}
                      </div>
                    </div>
                  </div>

                  {/* Settlement Pipeline */}
                  <div>
                    <div style={{ fontSize: '13px', fontWeight: 700, color: '#0F172A', marginBottom: '10px' }}>
                      Пошаговая цепочка расчетов в сделке:
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      {dealResult.steps.map((st, i) => (
                        <div
                          key={i}
                          style={{
                            display: 'flex',
                            alignItems: 'flex-start',
                            gap: '10px',
                            backgroundColor: '#F1F5F9',
                            padding: '10px 14px',
                            borderRadius: '6px',
                            fontSize: '13px',
                            color: '#1E293B'
                          }}
                        >
                          <span style={{
                            backgroundColor: '#DC2626',
                            color: '#FFFFFF',
                            borderRadius: '50%',
                            width: '20px',
                            height: '20px',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            fontSize: '11px',
                            fontWeight: 700,
                            flexShrink: 0
                          }}>
                            {i + 1}
                          </span>
                          <span>{st.replace(/^[0-9]\.\s*/, '')}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Recommended Contract Clause */}
                  <div style={{
                    backgroundColor: '#0F172A',
                    color: '#F8FAFC',
                    padding: '16px',
                    borderRadius: '8px',
                    fontSize: '12px',
                    lineHeight: 1.5
                  }}>
                    <div style={{ color: '#EF4444', fontWeight: 700, marginBottom: '6px', fontSize: '11px', letterSpacing: '0.5px' }}>
                      РЕКОМЕНДОВАННАЯ ФОРМУЛИРОВКА ДЛЯ ПДКП / ОДКП:
                    </div>
                    <em>{dealResult.recommendedContractClause}</em>
                  </div>

                </motion.div>

              </div>
            </motion.div>
          )}

          {/* TAB 2: FIGJAM SETTLEMENT FLOWCHART */}
          {activeTab === 'matrix' && (
            <motion.div key="matrix" {...fadeIn}>
              <div style={{
                backgroundColor: '#FFFFFF',
                borderRadius: '12px',
                padding: '24px',
                border: '1px solid #E2E8F0',
                boxShadow: '0 2px 4px rgba(0,0,0,0.03)'
              }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
                  <div>
                    <h2 style={{ fontSize: '18px', fontWeight: 700, margin: 0, color: '#0F172A' }}>
                      Карта финансовых сценариев расчетов (FigJam settlement tree)
                    </h2>
                    <p style={{ fontSize: '13px', color: '#64748B', margin: '2px 0 0 0' }}>
                      Точная оцифровка 8 ветвей взаиморасчетов доски FigJam
                    </p>
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button
                      onClick={() => setMatrixBranch('free')}
                      style={{
                        padding: '8px 14px',
                        borderRadius: '6px',
                        fontSize: '12px',
                        fontWeight: 700,
                        cursor: 'pointer',
                        border: 'none',
                        backgroundColor: matrixBranch === 'free' ? '#DC2626' : '#F1F5F9',
                        color: matrixBranch === 'free' ? '#FFFFFF' : '#475569'
                      }}
                    >
                      ВЕТВЬ I: Объект чистый
                    </button>
                    <button
                      onClick={() => setMatrixBranch('mortgaged')}
                      style={{
                        padding: '8px 14px',
                        borderRadius: '6px',
                        fontSize: '12px',
                        fontWeight: 700,
                        cursor: 'pointer',
                        border: 'none',
                        backgroundColor: matrixBranch === 'mortgaged' ? '#DC2626' : '#F1F5F9',
                        color: matrixBranch === 'mortgaged' ? '#FFFFFF' : '#475569'
                      }}
                    >
                      ВЕТВЬ II: Залог Продавца
                    </button>
                  </div>
                </div>

                {/* Grid of Nodes */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(290px, 1fr))', gap: '16px' }}>
                  {matrixBranch === 'free' ? (
                    <>
                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #16A34A'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#16A34A', textTransform: 'uppercase' }}>1.1. Наличные 100%</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Полностью наличные средства</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Задаток по ПДКП + остаток до ОДКП. Ст. 488 ГК РФ — без залога. Общая расписка на всю сумму в день ОДКП.
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #2563EB'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#2563EB', textTransform: 'uppercase' }}>1.2. СБР / Аккредитив</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Безопасные расчеты</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Задаток в день ПДКП. Остаток через номинальный счет или аккредитив в течение 1-5 дней после регистрации. В день сделки расписка ТОЛЬКО на задаток!
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #DC2626'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#DC2626', textTransform: 'uppercase' }}>1.3. Свои + МСК (Без ипотеки)</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Материнский капитал</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          <strong>Строгий ценз:</strong> ребенку &gt; 3 лет! Выплата из СФР до 21 дня после перехода права. Залог продавца по ст. 488 ГК РФ исключается в договоре.
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #7C3AED'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#7C3AED', textTransform: 'uppercase' }}>1.4. Ипотека Банка</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Целевой кредит банка</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Ипотека в силу закона ст. 77 102-ФЗ в пользу Банка. Статья 488 ГК РФ категорически не применяется.
                        </p>
                      </motion.div>
                    </>
                  ) : (
                    <>
                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #D97706'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#D97706', textTransform: 'uppercase' }}>2.1. До сделки наличными</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Гашение до сделки</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Покупатель дает деньги до ОДКП на снятие залога. Требует целевого займа и нотариальной доверенности на снятие залога.
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #DC2626'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#DC2626', textTransform: 'uppercase' }}>2.2. Трансферабельный аккредитив</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Авторасщепление сумм</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Золотой стандарт «Этажи»: банк автоматически направляет часть средств банку продавца на снятие залога, а дельту — продавцу.
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #2563EB'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#2563EB', textTransform: 'uppercase' }}>2.3. Ипотека на Ипотеку</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Внутри/межбанк</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          Внутрибанковская (Сбер на Сбер) единым днем. Межбанковская — под условие погашения первичного залога и наложения нового.
                        </p>
                      </motion.div>

                      <motion.div variants={cardVariant} initial="initial" animate="animate" whileHover="hover" style={{
                        border: '1px solid #E2E8F0', borderRadius: '10px', padding: '16px', backgroundColor: '#FFFFFF', borderTop: '4px solid #059669'
                      }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, color: '#059669', textTransform: 'uppercase' }}>2.4. Залог + Ипотека + МСК</div>
                        <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '6px 0' }}>Многоступенчатая схема</h3>
                        <p style={{ fontSize: '12px', color: '#475569', lineHeight: 1.4 }}>
                          ПВ покупателя частично гасит залог продавца, кредит покупателя закрывает остаток залога, МСК поступает продавцу после регистрации.
                        </p>
                      </motion.div>
                    </>
                  )}
                </div>

              </div>
            </motion.div>
          )}

          {/* TAB 3: CONTRACT BLOCKS */}
          {activeTab === 'contract_blocks' && (
            <motion.div key="contract_blocks" {...fadeIn}>
              <div style={{
                backgroundColor: '#FFFFFF',
                borderRadius: '12px',
                padding: '24px',
                border: '1px solid #E2E8F0',
                boxShadow: '0 2px 4px rgba(0,0,0,0.03)'
              }}>
                <h2 style={{ fontSize: '18px', fontWeight: 700, marginBottom: '6px', color: '#0F172A' }}>
                  Блочная структура предварительного договора (ПДКП)
                </h2>
                <p style={{ fontSize: '13px', color: '#64748B', marginBottom: '20px' }}>
                  Схема взаимосвязи разделов договора согласно FigJam (Welcome to FigJam)
                </p>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px' }}>
                  {[
                    {
                      title: '1. Преамбула и место заключения',
                      desc: 'г. Пенза, дата подписания, идентификация сделки и дееспособность сторон.'
                    },
                    {
                      title: '2. Стороны сделки',
                      desc: 'Продавец (взрослые, супруги, дети) и Покупатель. Проверка 152-ФЗ, СОПД, паспортных данных.'
                    },
                    {
                      title: '3. Предмет предварительного договора',
                      desc: 'Обязательство заключить в будущем Основной ДКП в согласованный срок в отношении объекта.'
                    },
                    {
                      title: '4. Описательные данные объекта',
                      desc: 'Кадастровый номер, точный адрес, площадь, назначение, этаж, правоустанавливающие документы.'
                    },
                    {
                      title: '5. Перечень мебели и техники',
                      desc: 'Исчерпывающий список передаваемого имущества (кухня, техника, мебель) без доплат.'
                    },
                    {
                      title: '6. Финансовые условия и порядок расчетов',
                      desc: 'Задаток, форма расчетов (СБР, аккредитив, нал), сроки, банковские реквизиты, статус залога.'
                    },
                    {
                      title: '7. Снятие с регистрационного учета',
                      desc: 'Срок выписки зарегистрированных лиц строго до заключения ОДКП или до передачи ключей.'
                    },
                    {
                      title: '8. Подписи и реквизиты сторон',
                      desc: 'Построчные подписи на всю ширину 17 см: Продавец, Покупатель, подтверждение получения задатка.'
                    }
                  ].map((block, idx) => (
                    <div
                      key={idx}
                      style={{
                        padding: '16px',
                        borderRadius: '8px',
                        border: '1px solid #E2E8F0',
                        backgroundColor: '#F8FAFC',
                        borderLeft: '4px solid #DC2626'
                      }}
                    >
                      <h4 style={{ fontSize: '14px', fontWeight: 700, margin: '0 0 6px 0', color: '#0F172A' }}>
                        {block.title}
                      </h4>
                      <p style={{ fontSize: '12px', color: '#475569', margin: 0, lineHeight: 1.4 }}>
                        {block.desc}
                      </p>
                    </div>
                  ))}
                </div>
              </div>
            </motion.div>
          )}

          {/* TAB 4: PREP CHECKLIST */}
          {activeTab === 'prep_checklist' && (
            <motion.div key="prep_checklist" {...fadeIn}>
              <div style={{
                backgroundColor: '#FFFFFF',
                borderRadius: '12px',
                padding: '24px',
                border: '1px solid #E2E8F0',
                boxShadow: '0 2px 4px rgba(0,0,0,0.03)'
              }}>
                <h2 style={{ fontSize: '18px', fontWeight: 700, marginBottom: '6px', color: '#0F172A' }}>
                  Чек-лист необходимой подготовки к сделке
                </h2>
                <p style={{ fontSize: '13px', color: '#64748B', marginBottom: '20px' }}>
                  Комплекты документов и регламент преддоговорного чекапа «Этажи»
                </p>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '20px' }}>
                  
                  {/* Seller docs */}
                  <div style={{
                    backgroundColor: '#F8FAFC',
                    padding: '18px',
                    borderRadius: '8px',
                    border: '1px solid #E2E8F0'
                  }}>
                    <h3 style={{ fontSize: '15px', fontWeight: 700, color: '#DC2626', marginBottom: '12px' }}>
                      📋 Документы от Продавца
                    </h3>
                    <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: '#334155', lineHeight: 1.6 }}>
                      <li>Паспорта всех собственников (свидетельства о рождении детей)</li>
                      <li>Правоустанавливающие документы (ДКП, приватизация, дарение, наследство)</li>
                      <li>Свежая выписка из ЕГРН (проверка арестов и запретов)</li>
                      <li>Нотариальное согласие супруга(и) либо брачный договор</li>
                      <li>Справка о зарегистрированных лицах (выписка из домовой книги)</li>
                      <li>Справка об отсутствии задолженности по ЖКУ и капремонту</li>
                      <li>При детях-собственниках: приказ Органов опеки и попечительства</li>
                    </ul>
                  </div>

                  {/* Buyer docs */}
                  <div style={{
                    backgroundColor: '#F8FAFC',
                    padding: '18px',
                    borderRadius: '8px',
                    border: '1px solid #E2E8F0'
                  }}>
                    <h3 style={{ fontSize: '15px', fontWeight: 700, color: '#2563EB', marginBottom: '12px' }}>
                      📋 Документы от Покупателя
                    </h3>
                    <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: '#334155', lineHeight: 1.6 }}>
                      <li>Паспорта всех покупателей / СНИЛС / ИНН</li>
                      <li>Нотариальное согласие супруга (при покупке без брачного договора)</li>
                      <li>При ипотеке: кредитный договор, график платежей, одобрение банка</li>
                      <li>При МСК: государственный сертификат, справка СФР об остатке средств</li>
                      <li>Платежное поручение / подтверждение размещения средств на СБР / аккредитиве</li>
                    </ul>
                  </div>

                  {/* Checks & Actions */}
                  <div style={{
                    backgroundColor: '#F8FAFC',
                    padding: '18px',
                    borderRadius: '8px',
                    border: '1px solid #E2E8F0'
                  }}>
                    <h3 style={{ fontSize: '15px', fontWeight: 700, color: '#059669', marginBottom: '12px' }}>
                      🔍 Обязательный юридический чекап
                    </h3>
                    <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: '#334155', lineHeight: 1.6 }}>
                      <li>Федресурс и КАД: проверка на банкротство Продавца и Покупателя</li>
                      <li>База ФССП: проверка наличия исполнительных производств и арестов</li>
                      <li>МВД: проверка действительности паспортов РФ</li>
                      <li>Росреестр: проверка соответствия кадастровых границ и планировки</li>
                      <li>Проверка истории перехода прав (защита от оспаривания сделок)</li>
                    </ul>
                  </div>

                </div>

              </div>
            </motion.div>
          )}

        </AnimatePresence>

      </div>
    </div>
  );
};

export default FigJamDealMatrix;
