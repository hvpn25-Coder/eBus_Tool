function derived = RCA_BuildDerivedSignals(signals, specs, referenceInfo, config)
% RCA_BuildDerivedSignals  Build aligned vehicle-level derived traces.

if nargin < 4 || isempty(config)
    config = RCA_Config();
end

t = referenceInfo.Time_s(:);
n = numel(t);
dt = [0; diff(t)];
dt(dt < 0) = 0;

derived = struct();
derived.time_s = t;
derived.dt_s = dt;

vehVel = localVector(signals, 'veh_vel', n);
vehPos = localVector(signals, 'veh_pos', n);
vehAcc = localVector(signals, 'veh_acc', n);
speedDemand = localVector(signals, 'veh_des_vel', n);
roadSlope = localVector(signals, 'road_slp', n);
ambientTemp = localVector(signals, 'amb_temp', n);
accPdl = localVector(signals, 'acc_pdl', n);
brkPdl = localVector(signals, 'brk_pdl', n);

if all(isnan(vehVel)) && ~all(isnan(vehPos))
    vehVel = [0; diff(vehPos)] .* 3600 ./ max(dt, eps);
end
if all(isnan(vehPos)) && ~all(isnan(vehVel))
    vehPos = cumtrapz(t, vehVel / 3600);
end
if all(isnan(vehAcc)) && ~all(isnan(vehVel))
    vehAcc = [0; diff(vehVel / 3.6)] ./ max(dt, eps);
end

emot1Trq = localVector(signals, 'emot1_act_trq', n);
emot2Trq = localVector(signals, 'emot2_act_trq', n);
emot1Dem = localVector(signals, 'emot1_dem_trq', n);
emot2Dem = localVector(signals, 'emot2_dem_trq', n);
emot1Max = localVector(signals, 'emot1_max_av_trq', n);
emot2Max = localVector(signals, 'emot2_max_av_trq', n);
emot1Min = localVector(signals, 'emot1_min_av_trq', n);
emot2Min = localVector(signals, 'emot2_min_av_trq', n);

emot1SpdRad = localVector(signals, 'emot1_act_spd', n);
emot2SpdRad = localVector(signals, 'emot2_act_spd', n);
emot1Pwr = localVector(signals, 'emot1_pwr', n);
emot2Pwr = localVector(signals, 'emot2_pwr', n);
emot1Loss = localVector(signals, 'emot1_loss_pwr', n);
emot2Loss = localVector(signals, 'emot2_loss_pwr', n);

grNum = localVector(signals, 'gr_num', n);
grRatio = localVector(signals, 'gr_ratio', n);
gbxOutTrq = localVector(signals, 'gbx_out_trq', n);
gbxOutSpd = localVector(signals, 'gbx_out_spd', n);
gbxLoss = localVector(signals, 'gbx_pwr_loss', n);

tractionForce = localVector(signals, 'net_trac_trq', n);
wheelForce = localVector(signals, 'whl_force', n);
if all(isnan(tractionForce)) && ~all(isnan(wheelForce))
    tractionForce = wheelForce;
end
if all(isnan(wheelForce))
    wheelForce = tractionForce;
end

fricBrakeForce = localVector(signals, 'fric_brk_force', n);
fricBrakePower = localVector(signals, 'fric_brk_pwr', n);
rollForce = localVector(signals, 'roll_res_force', n);
gradeForce = localVector(signals, 'grad_force', n);
aeroForce = localVector(signals, 'aero_drag_force', n);

battCurr = localVector(signals, 'batt_curr', n);
battVolt = localVector(signals, 'batt_volt', n);
battPwr = localVector(signals, 'batt_pwr', n);
battLoss = localVector(signals, 'batt_loss_pwr', n);
battSoc = localVector(signals, 'batt_soc', n);
battTemp = localVector(signals, 'batt_temp', n);

auxCurr = localVector(signals, 'aux_curr', n);
auxVolt = localVector(signals, 'aux_volt', n);
auxPwr = auxCurr .* auxVolt / 1000;

battChgCurrLim = localVector(signals, 'batt_chrg_curr_lim', n);
battDisCurrLim = localVector(signals, 'batt_dischrg_curr_lim', n);
battChgPwrLim = localVector(signals, 'batt_chrg_pwr_lim', n);
battDisPwrLim = localVector(signals, 'batt_dischrg_pwr_lim', n);

motorMechPwr = (emot1Trq .* emot1SpdRad + emot2Trq .* emot2SpdRad) / 1000;
motorElecPwr = localCombineSum(emot1Pwr, emot2Pwr);
motorLossPwr = localCombineSum(emot1Loss, emot2Loss);
motorSpeedRpm = localMeanAvailable(abs(emot1SpdRad) * 60 / (2 * pi), abs(emot2SpdRad) * 60 / (2 * pi));
torqueActual = localCombineSum(emot1Trq, emot2Trq);
torqueDemand = localCombineSum(emot1Dem, emot2Dem);
torquePositiveLimit = localCombineSum(emot1Max, emot2Max);
torqueNegativeLimit = localCombineSum(emot1Min, emot2Min);

vehVelMps = vehVel / 3.6;
tractionPower = tractionForce .* vehVelMps / 1000;
resistivePower = (max(rollForce, 0) + max(gradeForce, 0) + max(aeroForce, 0)) .* vehVelMps / 1000;
distanceStep = [0; diff(vehPos)];
if all(isnan(distanceStep))
    distanceStep = vehVel .* dt / 3600;
end

derived.referenceTimeSource = referenceInfo.Source;
derived.vehVel_kmh = vehVel;
derived.vehVel_mps = vehVelMps;
derived.vehiclePosition_km = vehPos;
derived.vehicleAcceleration_mps2 = vehAcc;
derived.speedDemand_kmh = speedDemand;
derived.speedError_kmh = speedDemand - vehVel;
derived.roadSlope_pct = roadSlope;
derived.ambientTemp_C = ambientTemp;
derived.accPedal_pct = accPdl;
derived.brkPedal_pct = brkPdl;

derived.torqueDemandTotal_Nm = torqueDemand;
derived.torqueActualTotal_Nm = torqueActual;
derived.torquePositiveLimit_Nm = torquePositiveLimit;
derived.torqueNegativeLimit_Nm = torqueNegativeLimit;
derived.motorSpeed_rpm = motorSpeedRpm;
derived.motorElectricalPower_kW = motorElecPwr;
derived.motorMechanicalPower_kW = motorMechPwr;
derived.motorLossPower_kW = motorLossPwr;

derived.gearNumber = grNum;
derived.gearRatio = grRatio;
derived.gearboxOutputTorque_Nm = gbxOutTrq;
derived.gearboxOutputSpeed_rads = gbxOutSpd;
derived.gearboxLossPower_kW = gbxLoss;

derived.tractionForce_N = tractionForce;
derived.wheelForce_N = wheelForce;
derived.tractionPower_kW = tractionPower;
derived.frictionBrakeForce_N = fricBrakeForce;
derived.frictionBrakePower_kW = fricBrakePower;

derived.rollingResistanceForce_N = rollForce;
derived.gradeForce_N = gradeForce;
derived.aeroDragForce_N = aeroForce;
derived.resistivePower_kW = resistivePower;

derived.batteryCurrent_A = battCurr;
derived.batteryVoltage_V = battVolt;
derived.batteryPower_kW = battPwr;
derived.batteryLossPower_kW = battLoss;
derived.batterySOC_pct = battSoc;
derived.batteryTemp_C = battTemp;

derived.battChargeCurrentLimit_A = battChgCurrLim;
derived.battDischargeCurrentLimit_A = battDisCurrLim;
derived.battChargePowerLimit_kW = battChgPwrLim;
derived.battDischargePowerLimit_kW = battDisPwrLim;

derived.auxiliaryCurrent_A = auxCurr;
derived.auxiliaryVoltage_V = auxVolt;
derived.auxiliaryPower_kW = auxPwr;

derived.distanceStep_km = distanceStep;
derived.tripDistance_km = max(vehPos, [], 'omitnan') - min(vehPos, [], 'omitnan');
if isnan(derived.tripDistance_km) || derived.tripDistance_km <= 0
    derived.tripDistance_km = sum(max(distanceStep, 0), 'omitnan');
end
derived.finalDriveRatio = localScalar(specs, 'FD_mec_iDiffAxle');
derived.vehicleMass_kg = localScalar(specs, 'VDy_mec_massVehicle_kg');
derived.wheelRadius_m = localScalar(specs, 'VDy_mec_rWheelDriven');

derived.motorSpeedRef_rpm = localSpecVector(specs, 'ED_emot_nTab_1stD');
derived.motorTorqueRef_Nm = localSpecVector(specs, 'ED_emot_trqMaxRow');
end

function vec = localVector(store, signalName, n)
signal = RCA_GetSignalData(store, signalName);
vec = NaN(n, 1);
if ~signal.Available
    return;
end
if ~isempty(signal.AlignedData)
    data = signal.AlignedData;
elseif ~isempty(signal.Data)
    data = signal.Data;
else
    return;
end
data = double(data(:));
if isscalar(data)
    vec = repmat(data, n, 1);
elseif numel(data) == n
    vec = data;
else
    sourceIndex = linspace(0, 1, numel(data));
    targetIndex = linspace(0, 1, n);
    vec = interp1(sourceIndex(:), data, targetIndex(:), 'linear', 'extrap');
end
end

function value = localScalar(store, specName)
spec = RCA_GetSignalData(store, specName);
if spec.Available && ~isempty(spec.Data)
    value = double(spec.Data(1));
else
    value = NaN;
end
end

function vec = localSpecVector(store, specName)
spec = RCA_GetSignalData(store, specName);
if spec.Available && ~isempty(spec.Data)
    vec = double(spec.Data(:));
else
    vec = NaN(0, 1);
end
end

function out = localCombineSum(a, b)
out = zeros(size(a));
out(:) = 0;
maskA = ~isnan(a);
maskB = ~isnan(b);
out(maskA) = out(maskA) + a(maskA);
out(maskB) = out(maskB) + b(maskB);
out(~maskA & ~maskB) = NaN;
end

function out = localMeanAvailable(a, b)
out = NaN(size(a));
maskA = ~isnan(a);
maskB = ~isnan(b);
out(maskA & ~maskB) = a(maskA & ~maskB);
out(~maskA & maskB) = b(~maskA & maskB);
out(maskA & maskB) = (a(maskA & maskB) + b(maskA & maskB)) / 2;
end
