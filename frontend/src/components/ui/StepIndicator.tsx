const STEPS = [
  "Upload",
  "Data Format",
  "Frequency",
  "Data Type",
  "Granularity",
  "Identifiers",
  "Generate",
];

const ARROW = 6; // px depth of the arrow point

function getClipPath(index: number, total: number) {
  const isFirst = index === 0;
  const isLast = index === total - 1;

  if (isFirst) {
    return `polygon(0 0, calc(100% - ${ARROW}px) 0, 100% 50%, calc(100% - ${ARROW}px) 100%, 0 100%)`;
  }
  if (isLast) {
    return `polygon(0 0, 100% 0, 100% 100%, 0 100%, ${ARROW}px 50%)`;
  }
  return `polygon(0 0, calc(100% - ${ARROW}px) 0, 100% 50%, calc(100% - ${ARROW}px) 100%, 0 100%, ${ARROW}px 50%)`;
}

interface Props {
  currentStep: number;
  onStepClick: (step: number) => void;
}

export default function StepIndicator({ currentStep, onStepClick }: Props) {
  return (
    <nav className="flex gap-0.5 mb-8">
      {STEPS.map((label, i) => {
        const step = i + 1;
        const isActive = step === currentStep;
        const isDone = step < currentStep;

        const bg = isActive
          ? "bg-teal-600"
          : isDone
          ? "bg-teal-50"
          : "bg-gray-200";

        const text = isActive
          ? "text-white"
          : isDone
          ? "text-teal-800"
          : "text-gray-500";

        const clip = getClipPath(i, STEPS.length);

        return (
          <div
            key={step}
            className="flex-1 bg-gray-400"
            style={{ clipPath: clip, padding: 0.5 }}
          >
            <button
              onClick={() => isDone && onStepClick(step)}
              disabled={!isDone}
              style={{
                clipPath: clip,
                paddingLeft: i === 0 ? 2 : ARROW + 2,
                paddingRight: i === STEPS.length - 1 ? 2 : ARROW + 2,
              }}
              className={`w-full flex items-center justify-center gap-1 py-1.5 sm:py-2 text-sm font-medium transition ${bg} ${text}
                ${isDone ? "cursor-pointer hover:brightness-95" : "cursor-default"}
              `}
            >
              <span className="text-xs font-bold shrink-0">
                {isDone ? "\u2713" : step}
              </span>
              <span className="hidden sm:inline text-xs whitespace-nowrap">{label}</span>
            </button>
          </div>
        );
      })}
    </nav>
  );
}
