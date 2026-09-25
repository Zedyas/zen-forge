import { CommandController } from './CommandController'
import { WindowScreen } from './WindowScreen'
import { TipProvider } from '../ui/Tip'
import { useUpdateEvents } from './updates'

export function App() {
  useUpdateEvents()
  return (
    <TipProvider>
      <WindowScreen />
      <CommandController />
    </TipProvider>
  )
}
