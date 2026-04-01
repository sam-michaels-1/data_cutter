import { useNavigate } from "react-router-dom";
import { useWizard } from "../hooks/useWizard";
import { useSession } from "../components/SessionProvider";
import StepIndicator from "../components/ui/StepIndicator";
import UploadStep from "../components/steps/UploadStep";
import InputFormatStep from "../components/steps/InputFormatStep";
import FrequencyStep from "../components/steps/FrequencyStep";
import DataTypeStep from "../components/steps/DataTypeStep";
import GranularityStep from "../components/steps/GranularityStep";
import IdentifiersStep from "../components/steps/IdentifiersStep";
import ReviewStep from "../components/steps/ReviewStep";

export default function ImportPage() {
  const { state, dispatch, nextStep, prevStep, goToStep } = useWizard();
  const { setSessionId } = useSession();
  const navigate = useNavigate();

  const canProceed = (() => {
    switch (state.currentStep) {
      case 1:
        return !!(state.sessionId && state.selectedSheet);
      case 2:
        // Format step: raw needs confirmed mapping, cleaned needs date columns + customer col
        if (state.inputFormat === "raw") {
          return !!state.confirmedMapping;
        }
        return state.dateColumns.length > 0 && !!state.customerNameCol;
      case 3:
        return !!state.dataFrequency;
      case 4:
        return !!state.dataType;
      case 5:
        return state.outputGranularities.length > 0;
      case 6:
        return true;
      case 7:
        return false;
      default:
        return false;
    }
  })();

  const handleViewDashboard = () => {
    if (state.downloadId) {
      setSessionId(state.downloadId);
      navigate("/dashboard");
    }
  };

  const renderStep = () => {
    switch (state.currentStep) {
      case 1:
        return <UploadStep state={state} dispatch={dispatch} />;
      case 2:
        return <InputFormatStep state={state} dispatch={dispatch} />;
      case 3:
        return <FrequencyStep state={state} dispatch={dispatch} />;
      case 4:
        return <DataTypeStep state={state} dispatch={dispatch} />;
      case 5:
        return <GranularityStep state={state} dispatch={dispatch} />;
      case 6:
        return <IdentifiersStep state={state} dispatch={dispatch} />;
      case 7:
        return (
          <ReviewStep
            state={state}
            dispatch={dispatch}
            onViewDashboard={handleViewDashboard}
          />
        );
      default:
        return null;
    }
  };

  return (
    <div className="max-w-2xl mx-auto px-4 py-8">
      <div className="text-center mb-6">
        <h1 className="text-2xl font-bold">Import Data</h1>
        <p className="text-gray-500 text-sm">
          Upload raw customer data and configure your analysis
        </p>
      </div>

      <StepIndicator currentStep={state.currentStep} onStepClick={goToStep} />

      {state.error && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-4 text-sm text-red-700">
          {state.error}
        </div>
      )}

      <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
        {renderStep()}
      </div>

      <div className="flex justify-between mt-6">
        <button
          onClick={prevStep}
          disabled={state.currentStep === 1}
          className={`px-5 py-2 rounded-lg font-medium transition
            ${state.currentStep === 1
              ? "text-gray-300 cursor-default"
              : "text-gray-600 hover:bg-gray-200"}
          `}
        >
          Back
        </button>

        {state.currentStep < 7 && (
          <button
            onClick={nextStep}
            disabled={!canProceed}
            className={`px-5 py-2 rounded-lg font-medium text-white transition
              ${canProceed
                ? "bg-teal-600 hover:bg-teal-700"
                : "bg-gray-300 cursor-not-allowed"}
            `}
          >
            Next
          </button>
        )}
      </div>
    </div>
  );
}
