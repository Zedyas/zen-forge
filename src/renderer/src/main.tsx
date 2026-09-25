import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { Toaster } from 'sonner'
import { App } from './app/App'
import './styles.css'

const root = document.getElementById('root')
if (root === null) throw new Error('The renderer root element is missing.')

createRoot(root).render(
  <StrictMode>
    <App />
    <Toaster
      position="bottom-right"
      offset={40}
      toastOptions={{ classNames: { toast: 'app-toast', description: 'app-toast-description' } }}
    />
  </StrictMode>,
)
