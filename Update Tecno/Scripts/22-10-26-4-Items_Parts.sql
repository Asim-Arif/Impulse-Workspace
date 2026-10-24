
GO

/****** Object:  Table [dbo].[Items_Parts]    Script Date: 10/22/2026 7:17:09 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Items_Parts](
	[EntryID] [int] IDENTITY(1,1) NOT NULL,
	[ItemID] [varchar](50) NOT NULL,
	[ItemID_Part] [varchar](50) NOT NULL,
	[Qty] [real] NULL,
	[ProcessID] [int] NULL,
 CONSTRAINT [PK_Items_Parts] PRIMARY KEY CLUSTERED 
(
	[ItemID] ASC,
	[ItemID_Part] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Items_Parts]  WITH NOCHECK ADD  CONSTRAINT [FK_Items_Parts_Items] FOREIGN KEY([ItemID])
REFERENCES [dbo].[Items] ([ItemID])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[Items_Parts] CHECK CONSTRAINT [FK_Items_Parts_Items]
GO

