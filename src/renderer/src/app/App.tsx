import { CommandController } from './CommandController'
import { WindowScreen } from './WindowScreen'
import { ConfirmHost } from '../ui/ConfirmDialog'
import { TipProvider } from '../ui/Tip'
import { useUpdateEvents } from './updates'

export function App() {
  useUpdateEvents()
  return (
    <TipProvider>
      <WindowScreen />
      <CommandController />
      {/* One host for the whole window, so a save asked from any tab, Home or a quit can show its question. */}
      <ConfirmHost />
    </TipProvider>
  )
}
