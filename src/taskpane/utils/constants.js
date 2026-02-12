/* Stier */
export const dokumenttypeurl = "https://localhost:3000/assets/";
export const tableStylesurl = "https://localhost:3000/assets/";
export const configurl = "https://raw.githubusercontent.com/Randers-Kommune-Digitalisering/budget-word-app-config/refs/heads/develop/";


/* Årstal */
const currentYear = new Date(Date.now()).getFullYear();
export const lastYear = currentYear - 1;
export const lastYear2 = currentYear - 2;
export const budgetperiodeÅr1 = currentYear + 1;
export const budgetperiodeÅr2 = currentYear + 2;
export const budgetperiodeÅr3 = currentYear + 3;
export const budgetperiodeÅr4 = currentYear + 4;
export const budgetperiode = budgetperiodeÅr1 + "-" + budgetperiodeÅr4;
