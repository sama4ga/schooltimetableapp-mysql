Public Class teacher
    Private teacher_name As String
    Private active_status As Boolean
    Private work_load As Integer
    Private teacher_id As Integer
    Private available_time As List(Of teachers_avialable_time)
    Private teacher_type As String

    ''' <summary>
    ''' Creates a new teacher
    ''' </summary>
    Public Sub New()
        teacher_name = "unnamed"
        'teacher_id = 0
        work_load = 0
        active_status = False
    End Sub

    ''' <summary>
    ''' Creates a new teacher with name
    ''' </summary>
    ''' <param name="name">Holds the name of the teacher</param>
    Public Sub New(name As String, type As String)
        teacher_name = name
        'teacher_id = 0
        work_load = 0
        active_status = False
        teacher_type = type
    End Sub

    ''' <summary>
    ''' Creates a new teacher with name and id
    ''' </summary>
    ''' <param name="name">Holds the name of the teacher</param>
    ''' <param name="id">Holds the id of the teacher</param>
    ''' <param name="type">Holds the type of the teacher</param>
    Public Sub New(name As String, id As Integer, type As String)
        teacher_name = name
        teacher_id = id
        work_load = 0
        active_status = False
        available_time = New List(Of teachers_avialable_time)
        teacher_type = type
    End Sub

    ''' <summary>
    ''' Holds the name of  the teacher
    ''' </summary>
    ''' <returns></returns>
    Public Property Name As String
        Get
            Return teacher_name
        End Get
        Set(value As String)
            teacher_name = value
        End Set
    End Property

    ''' <summary>
    ''' Holds the id of  the teacher
    ''' </summary>
    ''' <returns></returns>
    Public Property Id As Integer
        Get
            Return teacher_id
        End Get
        Set(value As Integer)
            teacher_id = value
        End Set
    End Property

    ''' <summary>
    ''' Used to hold the state of the teacher
    ''' </summary>
    ''' <returns>True if teacher is active and false otherwise</returns>
    Public Property Active As Boolean
        Get
            Return active_status
        End Get
        Set(value As Boolean)
            active_status = value
        End Set
    End Property

    ''' <summary>
    ''' Used to hold teacher's work load. Maximum is 4
    ''' </summary>
    ''' <returns>current teacher work load</returns>
    Public Property WorkLoad As Integer
        Get
            Return work_load
        End Get
        Set(value As Integer)
            work_load = value
        End Set
    End Property

    ''' <summary>
    ''' Holds the type of this teacher
    ''' </summary>
    ''' <returns></returns>
    Public Property Type As String
        Get
            Return teacher_type
        End Get
        Set(value As String)
            teacher_type = value
        End Set
    End Property

    ''' <summary>
    ''' Adds a new record of available time to the teacher
    ''' </summary>
    ''' <param name="time">The specfied start_time,end_time and day to be added</param>
    Public Sub AddAvailableTime(time As teachers_avialable_time)
        available_time.Add(time)
    End Sub

    ''' <summary>
    ''' Returns all the time teacher is available
    ''' </summary>
    ''' <returns>List of available teacher times</returns>
    Public ReadOnly Property AvailableTime As List(Of teachers_avialable_time)
        Get
            Return available_time
        End Get
    End Property

End Class
