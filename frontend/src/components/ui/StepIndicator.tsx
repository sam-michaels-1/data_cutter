const STEPS = [
  "Upload",
  "Data Format",
  "Frequency",
  "Data Type",
  "Granularity",
  "Identifiers",
  "Generate",
];

interface Props {
  currentStep: number;
  onStepClick: (step: number) => void;
}

export default function StepIndicator({ currentStep, onStepClick }: Props) {
  return (
    <nav className="flex items-center justify-center gap-2 mb-8">
      {STEPS.map((label, i) => {
        const step = i + 1;
        const isActive = step === currentStep;
        const isDone = step < currentStep;
        return (
          <button
            key={step}
            onClick={() => isDone && onStepClick(step)}
            disabled={!isDone}
            className={`flex items-center gap-2 px-3 py-1.5 rounded-full text-sm font-medium transition
              ${isActive ? "bg-teal-600 text-white" : ""}
              ${isDone ? "bg-teal-100 text-teal-700 cursor-pointer hover:bg-teal-200" : ""}
              ${!isActive && !isDone ? "bg-gray-100 text-gray-400 cursor-default" : ""}
            `}
          >
            <span
              className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold
              ${isActive ? "bg-white text-teal-600" : ""}
              ${isDone ? "bg-teal-600 text-white" : ""}
              ${!isActive && !isDone ? "bg-gray-300 text-white" : ""}
            `}
            >
              {isDone ? "\u2713" : step}
            </span>
            <span className="hidden sm:inline">{label}</span>
          </button>
        );
      })}
    </nav>
  );
}
