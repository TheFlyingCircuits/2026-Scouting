from dataclasses import dataclass, field

@dataclass
class SingleTeamSingleMatchEntry:
    team_num: int = 0
    qual_match_num: int = 0
    autoFuel: int = 0
    autoL1Climb: bool = False
    teleFuel: int = 0

    l1Climb: bool = False
    l2Climb: bool = False
    l3Climb: bool = False
    

    robotBroke: bool = False

    passes: float = 0.0
    auto: float = 0.0
    speed: float = 0.0
    pickupSpeed: float = 0.0
    scoringSpeed: float = 0.0
    driverDecisiveness: float = 0.0
    balance: float = 0.0
    wouldYouPick: float = 0.0

    defenseOnScoring: bool = False
    defenseOnIntaking: bool = False
    defenseOnPathing: bool = False

    commenter: str = ""
    comment: str = ""

@dataclass
class TeamData:
    team_num: int = 0
    match_data: list = field(default_factory=list)
    aveAutoPoints: float = 0.0
    aveAutoFuelPoints: float = 0.0
    aveAutoClimbPoints: float = 0.0

    aveTeleFuelPoint: float = 0.0
    aveTeleClimbPoints: float = 0.0
    aveTelePoints: float = 0.0
    avePoints: float = 0.0

    drivetrain: str = ""
    fuelCapacity: int = 0

    commenters = []
    comments = []

    drivetrain: str = ""
    swerve: bool = False
    riceScore: float = 0.0

    avePasses: float = 0.0
    robotBroke= [] 
    aveAuto: float = 0.0
    aveSpeed: float = 0.0
    avePickupSpeed: float = 0.0
    aveScoringSpeed: float = 0.0
    aveDriverDecisiveness: float = 0.0
    aveBalance: float = 0.0
    aveWouldYouPick: float = 0.0

@dataclass
class scoutingAccuracyMatch:
    overallInaccuracyRed: float = 0.0 
    autoInaccuracyRed: float = 0.0
    teleInaccuracyRed: float = 0.0
    endGameInaccuracyRed: float = 0.0
    allainceColorRed: str = ""
    matchNumRed: int = 0
    scouterOneNameRed: str = ""
    scouterTwoNameRed: str = ""
    scouterThreeNameRed: str = ""
    scouterOneInacuracyRed: float = 0.0
    scouterTwoInacuracyRed: float = 0.0
    scouterThreeInacuracyRed: float = 0.0
    scouterOneMissClimbRed: bool = False
    scouterTwoMissClimbRed: bool = False
    scouterThreeMissClimbRed: bool = False

    overallInaccuracyBlue: float = 0.0 
    autoInaccuracyBlue: float = 0.0
    teleInaccuracyBlue: float = 0.0
    endGameInaccuracyBlue: float = 0.0
    allainceColorBlue: str = ""
    matchNumBlue: int = 0
    scouterOneNameBlue: str = ""
    scouterTwoNameBlue: str = ""
    scouterThreeNameBlue: str = ""
    scouterOneInacuracyBlue: float = 0.0
    scouterTwoInacuracyBlue: float = 0.0
    scouterThreeInacuracyBlue: float = 0.0
    scouterOneMissClimbBlue: bool = False
    scouterTwoMissClimbBlue: bool = False
    scouterThreeMissClimbBlue: bool = False
    
