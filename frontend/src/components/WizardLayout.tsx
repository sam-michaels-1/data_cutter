import { useWizard } from "../hooks/useWizard";
import StepIndicator from "./ui/StepIndicator";
import UploadStep from "./steps/UploadStep";
import DataTypeStep from "./steps/DataTypeStep";
import GranularityStep from "./steps/GranularityStep";
import IdentifiersStep from "./steps/IdentifiersStep";
import ReviewStep from "./steps/ReviewStep";

export default function WizardLayout() {
  const { state, dispatch, nextStep, prevStep, goToStep } = useWizard();

  // Validation: can we proceed to the next step?
  const canProceed = (() => {
    switch (state.currentStep) {
      case 1:
        return !!(
          state.sessionId &&
          state.selectedSheet &&
          state.confirmedMapping
        );
      case 2:
        return !!state.dataType;
      case 3:
        return state.outputGranularities.length > 0;
      case 4:
        return true; // identifiers are optional
      case 5:
        return false; // no "next" on last step
      default:
        return false;
    }
  })();

  const renderStep = () => {
    switch (state.currentStep) {
      case 1:
        return <UploadStep state={state} dispatch={dispatch} />;
      case 2:
        return <DataTypeStep state={state} dispatch={dispatch} />;
      case 3:
        return <GranularityStep state={state} dispatch={dispatch} />;
      case 4:
        return <IdentifiersStep state={state} dispatch={dispatch} />;
      case 5:
        return <ReviewStep state={state} dispatch={dispatch} />;
      default:
        return null;
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-2xl mx-auto px-4 py-8">
        {/* Header */}
        <div className="text-center mb-6">
          <h1 className="text-2xl font-bold text-gray-900">
            Data Cutter
          </h1>
          <p className="text-gray-500 text-sm">
            Transform raw customer data into formatted analysis workbooks
          </p>
        </div>

        <StepIndicator
          currentStep={state.currentStep}
          onStepClick={goToStep}
        />

        {/* Error banner */}
        {state.error && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">
            {state.error}
          </div>
        )}

        {/* Step content */}
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
          {renderStep()}
        </div>

        {/* Navigation */}
        <div className="flex justify-between mt-6">
          <button
            onClick={prevStep}
            disabled={state.currentStep === 1}
            className={`px-5 py-2 rounded-lg font-medium transition
              ${state.currentStep === 1 ? "text-gray-300 cursor-default" : "text-gray-600 hover:bg-gray-200"}
            `}
          >
            Back
          </button>

          {state.currentStep < 5 && (
            <button
              onClick={nextStep}
              disabled={!canProceed}
              className={`px-5 py-2 rounded-lg font-medium text-white transition
                ${canProceed ? "bg-blue-600 hover:bg-blue-700" : "bg-gray-300 cursor-not-allowed"}
              `}
            >
              Next
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
