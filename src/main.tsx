// 外部通信ガードは他の何よりも先に有効化する
import { installNetworkGuard } from './security/networkGuard';
installNetworkGuard();

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

createRoot(document.getElementById('root') as HTMLElement).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
