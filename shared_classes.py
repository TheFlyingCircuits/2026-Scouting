from dataclasses import dataclass, field

@dataclass
class SingleTeamSingleMatchEntry:
    team_num: int = 0
    qual_match_num: int = 0
    autoFuel: int = 0
    autoPasses: int = 0
    autoL1Climb: boolean = False
    teleFuel: int = 0
    telePasses: int = 0

    l1Climb: boolean = False
    l2Climb: boolean = False
    l3Climb: boolean = False
    
    commenter: str = ""
    comment: str = ""
    totalPoints: int  = 0
    robotBroke: str = ""
    auto: str = ""
    speed: str = ""
    pickupSpeed: str = ""
    scoring: str = ""
    driverDecisiveness: str = ""
    balance: str = ""
    wouldYouPick: str = ""

@dataclass
class TeamData:
    team_num: int = 0
    match_data: list = field(default_factory=list)
    aveAutoPoints: float = 0
    aveLeavePoints: float = 0
    aveAutoL4Points: float = 0
    aveAutoL3Points: float = 0
    aveAutoL2Points: float = 0
    aveAutoL1Points: float = 0
    aveAutoProcessorPoints: float = 0
    aveAutoNetPoints: float = 0
    aveTelePoints: float = 0
    aveTeleL4Points: float = 0
    aveTeleL3Points: float = 0
    aveTeleL2Points: float = 0
    aveTeleL1Points: float = 0
    aveCoralPoints: float = 0
    aveTeleProcessorPoints: float = 0
    aveTeleNetPoints: float = 0
    aveBargePoints: float = 0
    aveAlgaePoints: float = 0
    avePoints: float = 0
    weight: float = 0
    drivetrain: str = ""
    whatGamePeiceCanIntake: str = ""
    howDoesScoreCoral: str = ""
    howDoesScoreAlgae: str = ""
    canClimbFrom: str = ""
    syle: str = ""
    aveAlgaeRemoved: float = 0
    noClimb: int = 0
    park: int = 0
    shallowClimb: int = 0
    deepClimb: int = 0
    commentNum: int = 1
    quantativeAve: float = 1
    drivetrain: str = ""
    aveSpeed: float = 0
    aveDriver: float = 0
    swerve: str = ""
    coral: str = ""
    algae: str = ""
    climb: str  = "" 
    riceScore: float = 0

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
    
