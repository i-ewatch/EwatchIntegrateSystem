﻿CREATE TABLE [dbo].[ElectricWeb]
(
	[ttime] NVARCHAR(14) NOT NULL,
    [ttimen] DATETIME NOT NULL,
	[CardNumber] NVARCHAR(6) NOT NULL DEFAULT 000000,
    [BordNumber] NVARCHAR(2) NOT NULL DEFAULT 00,
    [DeviceNumber] INT NOT NULL DEFAULT 0,
    [Loop] INT NOT NULL DEFAULT 0,
    [Va] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vnavg] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vab] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vbc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vca] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Vlavg] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Ia] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Ib] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Ic] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Iavg] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Freq] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Kwa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Kwb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Kwc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [Kwtot] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVARa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVARb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVARc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVARtot] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVAa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVAb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVAc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KVAtot] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PFa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PFb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PFc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PFavg] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleVa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleVb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleVc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleIa] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleIb] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [PhaseAngleIc] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    [KWH] DECIMAL(18, 2) NOT NULL DEFAULT 0,
    PRIMARY KEY ([CardNumber], [BordNumber], [DeviceNumber], [Loop]), 
)