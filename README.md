# ControlCard

Набор макросов Excel и вспомогательных Lua-скриптов для подготовки расчётных таблиц на основе данных, выгруженных из Archicad/ArchiFrame.

## Структура репозитория

- `excel/ControlCardTemplateEng.xlsm` — шаблон рабочей книги.
- `vba/` — исходные модули макросов VBA.
- `lua/ArchiFrameListing.lua` — скрипт экспорта данных из ArchiFrame в Excel.

## Основные макросы

- **Cutting Planks.vba** — создаёт лист «Раскрой Древесины» из «ИсходныеДанные» с учётом параметров.
- **CopyWoodQuantity.vba** — переносит объёмы древесины в лист «Вспомогательная (Панели)».
- **Cutting Plan Sheets2d.vba** — строит карты раскроя плит и формирует лист «Раскрой Плит».
- **CalculationSheet.vba** — собирает итоговый «Расчётный лист» с подсчётом объёма и массы.
- **PaintingSheet.vba** — формирует отчёт по покраске и отгрузке на основе листов раскроя.
- **Module6.vba** — служебные процедуры для справочных листов.

## Как использовать

1. С помощью `lua/ArchiFrameListing.lua` выгрузите элементы из ArchiFrame в Excel.
2. В открывшемся Excel файле поочерёдно запустите макросы с помощью кнопок на листе «ИсходныеДанные» или из меню «Разработчик»:
   - `GenerateCuttingPlan`;
   - `CopyWoodQuantity`;
   - `CuttingPlanSheets2D`;
   - `CreateCalculationSheet`;
   - `GeneratePaintingShippingReport_Full`;
   - при необходимости `WriteReferenceFormulas`.

После выполнения макросов книга будет содержать листы с раскроем, ведомостями материалов и сводными расчётами.