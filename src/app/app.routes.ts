import { Routes } from '@angular/router';
import { CotizadorComponent } from './cotizador/cotizador.component';

export const routes: Routes = [
  { path: 'cotizador', component: CotizadorComponent },
  { path: '', redirectTo: '/cotizador', pathMatch: 'full' }
];

