#RUN FILE

# Calls the Packages used for the optimization problem
using JuMP
using Printf
using Gurobi
#using CPLEX
using MathOptInterface
using JLD
using TimerOutputs
using DataFrames
using XLSX
using Parameters
using Dates
using CSV

# Calls the other Julia files
include("Structures.jl")
include("SetInputParameters.jl")
include("solveOptimizationAlgorithm.jl")        #solveOptimizationAlgorithm_3cuts
include("ProblemFormulationInequalities.jl")      #ProblemFormulationCutsTaylor_3
include("Saving in xlsx.jl")

date = string(today())

# PREPARE INPUT DATA
to = TimerOutput()

@timeit to "Set input data" begin

  #Set run case - indirizzi delle cartelle di input ed output
  case = set_runCase()
  @unpack (DataPath,InputPath,ResultPath,CaseName) = case;

  # Set run mode (how and what to run) and Input parameters
  runMode = read_runMode_file()
  InputParameters = set_parameters(runMode, case)
  @unpack (NYears, NMonths, NHoursStep, NHoursStage, NStages, NSteps, Big, conv, disc)= InputParameters;

  # Set solver parameters (Gurobi etc)
  SolverParameters = set_solverParameters()

  # Read power prices from a file [€/MWh]
  
  Battery_price = read_csv("Battery_decreasing_prices_mid.csv",case.DataPath)

  Pp14 = read_csv("prices_2014_8760.csv", case.DataPath);
  Pp15 = read_csv("prices_2015_8760.csv", case.DataPath);
  Pp16 = read_csv("prices_2016_8760.csv", case.DataPath);
  Pp17 = read_csv("prices_2017_8760.csv", case.DataPath);
  Pp18 = read_csv("prices_2018_8760.csv", case.DataPath);
  Pp19 = read_csv("prices_2019_8760.csv", case.DataPath);
  Pp20 = read_csv("prices_2020_8760.csv", case.DataPath);
  Pp21 = read_csv("prices_2021_8760.csv", case.DataPath);
  Pp22 = read_csv("prices_2022_8760_new.csv", case.DataPath);
  Pp23 = read_csv("prices_2023_8760.csv", case.DataPath);

  Power_prices = vcat(Pp14,Pp15,Pp16,Pp17,Pp18,Pp19,Pp20,Pp21,Pp22,Pp23);    
  
  # Upload battery's characteristics
  Battery = set_battery_system(runMode, case)
  @unpack (min_SOC, max_SOC, Eff_charge, Eff_discharge, min_P, max_P, max_SOH, min_SOH, Nfull) = Battery; 


  # Where and how to save the results
  FinalResPath= set_run_name(case, ResultPath, InputParameters)

end

#save input data
@timeit to "Save input" begin
    save(joinpath(FinalResPath,"CaseDetails.jld"), "case" ,case)
    save(joinpath(FinalResPath,"SolverParameters.jld"), "SolverParameters" ,SolverParameters)
    save(joinpath(FinalResPath,"InputParameters.jld"), "InputParameters" ,InputParameters)
    save(joinpath(FinalResPath,"BatteryCharacteristics.jld"), "BatteryCharacteristics" ,Battery)
    save(joinpath(FinalResPath,"PowerPrices.jld"),"PowerPrices",Power_prices)
end

@timeit to "Solve optimization problem" begin
  ResultsOpt = solveOptimizationProblem(InputParameters,SolverParameters,Battery);
  save(joinpath(FinalResPath, "optimization_results.jld"), "optimization_results", ResultsOpt)
end

# SAVE DATA IN EXCEL FILES
if runMode.excel_savings
  cartella = "C:\\GitSource-Batteries\\Batteries - exact - inequalities\\RESULTS"
  cd(cartella)
  Saving = data_saving(InputParameters,ResultsOpt)
else
  println("Solved without saving results in xlsx format.")
end


#end

print(to)




