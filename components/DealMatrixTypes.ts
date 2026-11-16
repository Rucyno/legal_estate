/**
 * Типы данных для интерактивной карты сделок и матрицы расчетов FigJam
 * Корпоративный стандарт «Этажи»
 */

export type PropertyType = 
  | 'apartment'           // Квартира (целиком)
  | 'apartment_share'     // Доля в квартире
  | 'house_and_land'      // Жилой дом и земельный участок
  | 'land'                // Земельный участок
  | 'room'                // Комната (с уведомлением соседей по ст. 250 ГК РФ)
  | 'commercial'          // Нежилое коммерческое помещение
  | 'commercial_and_land' // Нежилое коммерческое помещение + участок
  | 'land_and_outbuilding'; // Земельный участок и нежилой объект (баня, гараж, сарай)

export type PartyComposition = 
  | 'adult_1'             // 1 взрослый
  | 'adult_2'             // 2 взрослых
  | 'adult_3'             // 3 взрослых
  | 'spouses'             // 2 взрослых (супруги)
  | 'minor_under_14'      // Несовершеннолетний до 14 лет (родители)
  | 'minor_14_to_18'      // Несовершеннолетний от 14 до 18 лет (лично с согласия)
  | 'combined_family';    // Семья с детьми

export type EncumbranceStatus = 
  | 'free'                // Объект чистый (без залога и обременений)
  | 'mortgaged';          // Объект в залоге кредитора (ипотека Продавца)

export type PaymentMethod = 
  | 'cash_full'           // Наличные полностью
  | 'sbr'                 // СБР (система безопасных расчетов)
  | 'letter_of_credit'    // Аккредитив (покрытый безотзывный / трансферабельный)
  | 'safe_deposit_box';   // Банковская сейфовая ячейка

export type FinancingSource = 
  | 'own_funds'           // 100% собственные средства
  | 'own_plus_matcap'     // Собственные + МСК
  | 'own_plus_mortgage'   // Собственные + Ипотека банка
  | 'combined_all';       // Собственные + Ипотека + МСК

export interface DealScenarioResult {
  title: string;
  encumbranceRule: string;
  statuteArticle: string;
  matcapRule?: string;
  receiptRule: string;
  steps: string[];
  risksAndWarnings: string[];
  recommendedContractClause: string;
}
